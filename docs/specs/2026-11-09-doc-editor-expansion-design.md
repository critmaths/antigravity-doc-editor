# Design Specification: Antigravity Document Editor (`doc-editor`)

**Date:** 2026-11-09  
**Status:** Pending Approval  
**Project / Repo Name:** `antigravity-doc-editor`  
**Plugin Identifier:** `doc-editor`  
**Location:** `~/.gemini/config/plugins/doc-editor/` (installed) & `C:\Users\critm\.gemini\antigravity\scratch\antigravity-doc-editor\` (source repo)

---

## 1. Intent & Purpose
In Antigravity 2.0, opening either Microsoft Word (`.docx`) or OpenDocument Text (`.odt`) files displays *"Preview is not available for this file type"*. Users are forced to exit Antigravity and launch external desktop office suites just to review project requirements, outlines, or meeting notes.

This project expands our proven, zero-dependency plugin into **Antigravity Document Editor** (`doc-editor`): a unified, in-IDE word processor operating in Antigravity's auxiliary side pane. It delivers visual document previewing, rich text formatting, bidirectional disk synchronization for both `.docx` and `.odt`, and cross-format conversion.

---

## 2. Goals, Non-Goals, and Anti-Goals

### Goals
- **Dual-Engine Office Support**:
  - Full `.odt` (OASIS OpenDocument Text via `content.xml`) parsing and serialization.
  - Full `.docx` (ECMA-376 Office Open XML via `word/document.xml` & `word/styles.xml`) parsing and serialization.
- **In-IDE Auxiliary Sidecar**:
  - Live side-by-side editing in the auxiliary pane (`sidecar://doc-editor/doc_editor/ui.html`).
  - Dropdown workspace switcher listing both `.docx` and `.odt` files.
- **Rich Text WYSIWYG Formatting**:
  - Typography: Fonts (Inter, Arial, Calibri, Times New Roman, Georgia, Monospace), font sizes (10pt–32pt).
  - Styles: Bold, Italic, Underline, Strikethrough, Text color, Highlight color.
  - Structure: Headings (H1–H3), Paragraphs, Blockquotes, Tables.
  - Alignment: Left, Center, Right, Justify.
  - Lists: Bulleted and Numbered lists.
  - Actions: Undo/Redo, Dirty-state indicator, Ctrl+S keyboard shortcut.
- **Cross-Format Conversion & Export**:
  - One-click *"Save Copy As .docx"* when working in `.odt`.
  - One-click *"Save Copy As .odt"* when working in `.docx`.
  - Export to clean Markdown.
- **Zero External Dependencies**:
  - Runs natively on Windows using PowerShell and built-in .NET `System.Net.HttpListener`, `System.IO.Compression.ZipFile`, and XML DOM. No Node.js or Python required.
- **Agent Skill & Integration**:
  - Rules to automatically surface the sidecar editor pill link when `.docx` or `.odt` files are referenced in chat.
  - CLI converter mode for the agent to inspect and summarize documents.

### Non-Goals
- Complex drawing canvas shapes (WordArt, 3D charts, SmartArt).
- Multi-column complex newsletter layouts and print-press pagination.
- Running inside the native Monaco code tab (which is restricted in core Electron).

### Anti-Goals
- No lossy format flattening (do not convert to plain text and throw away document structures).
- No external cloud services or telemetry; document processing remains strictly local on the user's machine.

---

## 3. Architecture & Technical Design

### A. Repository & Plugin Structure
```text
antigravity-doc-editor/
├── plugin.json                 # Manifest: name="doc-editor", displayName="Antigravity Document Editor"
├── README.md                   # Public showcase documentation & quick start
├── LICENSE                     # MIT License
├── rules/
│   └── AGENTS.md               # Context rules: triggers on *.odt and *.docx files
├── skills/
│   └── doc-editor/
│       └── SKILL.md            # Agent skill: instructions, CLI usage, and sidecar links
├── sidecars/
│   └── doc_editor/
│       ├── sidecar.json        # Sidecar definition (powershell host, port, views)
│       ├── server.ps1          # HTTP server serving UI and REST APIs
│       ├── doc_converter.ps1   # Unified converter engine (ODT + DOCX bidirectional)
│       ├── ui.html             # Document editor webview
│       ├── editor.css          # Styling (paper canvas, toolbar, Antigravity theme)
│       └── editor.js           # Client WYSIWYG logic, format switcher, and save actions
└── tests/
    ├── create_sample.ps1       # Generates generic sample .odt and .docx files
    ├── test_converter.ps1      # ODT & DOCX conversion tests
    ├── test_server.ps1         # HTTP API endpoint tests
    ├── test_ui_assets.ps1      # Frontend asset tests
    ├── test_plugin_manifest.ps1# Plugin schema & sidecar configuration tests
    └── test_e2e.ps1            # End-to-end load, edit, save, and conversion tests
```

### B. DOCX Architecture & Office Open XML Serialization
A `.docx` file is a ZIP archive containing:
- `[Content_Types].xml`: Declares MIME types for parts.
- `_rels/.rels`: Top-level package relationships.
- `word/document.xml`: Main document body containing `<w:body>`, `<w:p>`, `<w:r>`, `<w:t>`, and trailing `<w:sectPr>`.
- `word/styles.xml`: Document paragraph and run styles (`Heading1`, `Heading2`, `Heading3`, `Normal`, `Quote`, etc.).
- `word/numbering.xml`: Numbering definitions for bulleted (`numId="1"`) and numbered (`numId="2"`) lists.
- `word/_rels/document.xml.rels`: Document-level relationships (styles, numbering, external hyperlinks).

**Builder Call — Non-Corrupting OpenXML Preservation**:
- When saving to an existing `.docx`, the serializer preserves any trailing `<w:sectPr>` node inside `<w:body>` so page setup, margins, and header/footer references remain intact and Microsoft Word never shows a file recovery prompt.
- When saving lists (`<ul>`, `<ol>`), the serializer ensures `word/numbering.xml`, its `[Content_Types].xml` override, and its `word/_rels/document.xml.rels` relationship exist.
- When saving hyperlinks (`<a href="...">`), the serializer registers an external `Relationship` in `word/_rels/document.xml.rels` and emits `<w:hyperlink r:id="rId...">`.

**Converter Routing (`doc_converter.ps1`):**
- **Reading (`Convert-DocumentToHtml`)**:
  - `.odt` $\rightarrow$ parses `<office:text>`, `<text:p>`, `<text:h>`, `<text:span>`, `<text:list>`, `<table:table>`, `<text:a>` $\rightarrow$ HTML.
  - `.docx` $\rightarrow$ parses `<w:p>`, `<w:pPr>` (headings, alignment, list `<w:numPr>`, quote style), `<w:r>`, `<w:rPr>` (bold `<w:b/>`, italic `<w:i/>`, underline `<w:u/>`, strike `<w:strike/>`, color `<w:color/>`, highlight/shading `<w:highlight/>` / `<w:shd/>`, font `<w:rFonts/>`, size `<w:sz/>`), `<w:hyperlink>`, `<w:tbl>` $\rightarrow$ HTML.
- **Writing (`Convert-HtmlToDocument`)**:
  - `.odt` $\rightarrow$ serializes HTML to `<office:text>` and updates `content.xml`.
  - `.docx` $\rightarrow$ serializes HTML to `<w:body>` (retaining `<w:sectPr>`) and updates `word/document.xml`, `word/numbering.xml`, and `word/_rels/document.xml.rels`.
- **Cross-Format Conversion & Export (`Convert-DocumentFormat`)**:
  - Supports converting `.odt` $\leftrightarrow$ `.docx` and exporting to `.md` (Markdown).
  - Creates a clean, standards-compliant target container from scratch when the target file does not yet exist, or updates it when overwriting.

### C. Server Endpoints (`server.ps1`)
- `GET /api/documents`: Returns list of all `.odt` and `.docx` files in the workspace.
- `GET /api/document?path=...`: Returns JSON with document title, format (`odt` or `docx`), and parsed HTML.
- `POST /api/document`: Saves updated HTML back to the active `.odt` or `.docx` file on disk (or creates a new `.odt`/`.docx` file if creating a new document).
- `POST /api/convert`: Accepts `{ sourcePath, targetPath, html, format }` (`docx`, `odt`, or `md`) to convert/export the current editor state directly to `targetPath` on disk.

### D. Frontend Toolbar, Modals & Format Switcher (`ui.html` & `editor.js`)
- **Builder Call — Zero Native `prompt()`/`confirm()` Dependence**: Because Electron auxiliary webviews often block native `window.prompt()` dialogs, all interactive prompts (Save As / Convert path confirmation, New Document creation, Insert Link URL, Insert Table dimensions, Unsaved Changes warning) use a built-in, keyboard-accessible HTML modal overlay inside `ui.html`.
- **Document Header & Metadata**:
  - Displays active document title, format badge (`ODT` or `DOCX`), save status indicator (`Saved` vs `• Unsaved changes`), workspace file switcher dropdown, **New** button, **Save** button (Ctrl+S), and **Save As / Export** menu.
- **Complete Interactive Toolbar (Every Button Wired & Verified)**:
  - Block style selector (`Normal Text`, `Heading 1`, `Heading 2`, `Heading 3`, `Quote`)
  - Font family selector (`Inter`, `Calibri`, `Arial`, `Times New Roman`, `Georgia`, `Monospace`)
  - Font size selector (`10pt`, `12pt`, `14pt`, `18pt`, `24pt`, `32pt`)
  - Inline formatting toggles (`Bold`, `Italic`, `Underline`, `Strikethrough`) with active state highlighting
  - Color pickers (`Text Color` and `Highlight Color`) with live visual color bar indicator
  - Alignment buttons (`Align Left`, `Align Center`, `Align Right`, `Justify`)
  - List buttons (`Bulleted List`, `Numbered List`)
  - Insert tools (`Insert Link` modal, `Insert Table` modal, `Clear Formatting`)
  - History buttons (`Undo`, `Redo`)
- **Save As / Export Menu**:
  - *Save Copy as DOCX (`.docx`)*: Opens modal pre-filled with `<current_base>.docx`, writes converted `.docx` to disk, updates file list, and switches to or confirms the saved file.
  - *Save Copy as ODT (`.odt`)*: Opens modal pre-filled with `<current_base>.odt`, writes converted `.odt` to disk, updates file list.
  - *Export as Markdown (`.md`)*: Opens modal pre-filled with `<current_base>.md`, writes clean Markdown to disk.

### E. Migration Strategy (`odt-editor` $\rightarrow$ `doc-editor`)
- **Builder Call — Clean In-Place Upgrade**:
  - Retire the standalone `odt-editor` plugin directory (`~/.gemini/config/plugins/odt-editor`) and replace it with `~/.gemini/config/plugins/doc-editor`.
  - Rename the GitHub repository `critmaths/antigravity-odt-editor` to `critmaths/antigravity-doc-editor` (GitHub automatically redirects old links) and move local source to `C:\Users\critm\.gemini\antigravity\scratch\antigravity-doc-editor`.

---

## 4. Verification & Testing Strategy
1. **Generic Sample Generation**: `create_sample.ps1` generates both a clean `sample.odt` and `sample.docx` containing standard headings, styled text, and bullet lists.
2. **DOCX Parsing & Round-Trip Test**: Verify `sample.docx` loads into HTML, receives edits, and serializes back into valid `word/document.xml`.
3. **ODT Parsing & Round-Trip Test**: Verify existing ODT regression tests remain 100% passing.
4. **Cross-Conversion Test**: Convert `sample.odt` to `sample.docx`, verify Word structure; convert `sample.docx` to `sample.odt`, verify ODF structure.
5. **Full E2E Test**: Run `test_e2e.ps1` testing both formats against live HTTP sidecar endpoints.
