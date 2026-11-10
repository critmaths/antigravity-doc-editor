# Antigravity Document Editor (`doc-editor`) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expand the plugin into **Antigravity Document Editor** (`doc-editor` / `antigravity-doc-editor`), supporting full in-IDE previewing, WYSIWYG editing, live saving, and cross-format conversion for both `.docx` (Microsoft Word) and `.odt` (OpenDocument Text) files, with every toolbar button, menu item, link, and modal dialog verified end-to-end.

**Architecture:** A unified PowerShell converter (`doc_converter.ps1`) handles bidirectional conversion between HTML and both OASIS ODF (`.odt`) and ECMA-376 OpenXML (`.docx`), plus Markdown export. A local `.NET HttpListener` server (`server.ps1`) exposes document discovery, load, save, creation, and cross-format conversion endpoints. The auxiliary pane frontend (`ui.html`, `editor.css`, `editor.js`) provides a complete toolbar, custom in-webview modals (avoiding blocked `window.prompt()`), format badges, and live disk sync.

**Tech Stack:** PowerShell 5.1+, .NET `System.Net.HttpListener`, `System.IO.Compression.ZipFile`, `System.Xml.XmlDocument`, Vanilla HTML5/CSS3/JS.

**Spec:** [2026-11-09-doc-editor-expansion-design.md](file:///C:/Users/critm/.gemini/antigravity/scratch/antigravity-odt-editor/docs/specs/2026-11-09-doc-editor-expansion-design.md)

## Global Constraints
- Target installed plugin directory: `C:\Users\critm\.gemini\config\plugins\doc-editor`
- Target source repository directory: `C:\Users\critm\.gemini\antigravity\scratch\antigravity-doc-editor`
- Zero external runtime dependencies (no Node.js, Python, or third-party binaries).
- Zero hardcoded user paths or personal document names in repository files or tests.
- Every `.SetAttribute()` and DOM mutation in PowerShell must suppress return values (`[void]`) so pipeline output is never polluted.
- Zero reliance on native `window.prompt()` or `window.confirm()` in `editor.js`; use the built-in HTML modal system.

## Review Focus
1. **OpenXML (`.docx`) package validity**: Saved and newly created `.docx` files must include valid `[Content_Types].xml`, `_rels/.rels`, `word/document.xml` (preserving `<w:sectPr>`), `word/styles.xml`, `word/numbering.xml`, and `word/_rels/document.xml.rels` so Word/LibreOffice never report corruption.
2. **All formatting round-trips across both `.docx` and `.odt`**: Headings (`h1`–`h3`), bold, italic, underline, strikethrough, text color, highlight color, font family, font size, alignment (`left`, `center`, `right`, `justify`), blockquotes, bullet/numbered lists, links (`<a href>`), and tables (`<table>`).
3. **Cross-format conversion**: `.odt` $\rightarrow$ `.docx`, `.docx` $\rightarrow$ `.odt`, and export to `.md` must write valid target files on disk.
4. **All UI controls & modals**: Every toolbar button, dropdown, link/table insertion modal, New Document modal, and Save As / Export menu item must be wired and tested.
5. **Path validation & security**: `server.ps1` must validate `.odt` / `.docx` / `.md` extensions and handle paths with spaces properly.

---

### Task 1: Dual-Format Sample Generator & Unified Converter Engine (`create_sample.ps1` & `doc_converter.ps1`)

**Files:**
- Modify: `tests/create_sample.ps1`
- Create: `sidecars/doc_editor/doc_converter.ps1`
- Modify: `tests/test_converter.ps1`

**Interfaces:**
- Produces in `tests/create_sample.ps1`:
  - `New-SampleOdt([string]$OutputPath)`
  - `New-SampleDocx([string]$OutputPath)`
- Produces in `sidecars/doc_editor/doc_converter.ps1`:
  - `New-BlankDocument([string]$DocPath)` (creates valid blank `.odt` or `.docx`)
  - `Convert-DocumentToHtml([string]$DocPath) -> [hashtable]@{ title = [string]; format = [string]; html = [string]; styles = [string] }`
  - `Convert-HtmlToDocument([string]$DocPath, [string]$HtmlContent) -> [bool]`
  - `Convert-DocumentFormat([string]$SourcePath, [string]$TargetPath, [string]$HtmlContent) -> [bool]`
  - `Convert-HtmlToMarkdown([string]$HtmlContent) -> [string]`
  - CLI entrypoint: `-Read <path> [-Format html|markdown|text]`

- [ ] **Step 1: Write failing tests in `tests/test_converter.ps1`**
  Update `tests/test_converter.ps1` to generate both `sample.odt` and `sample.docx`, test `Convert-DocumentToHtml` on both, test `Convert-HtmlToDocument` with rich HTML (headings, bold, italic, underline, strike, color, highlight, font, size, alignment, quote, ul/ol lists, hyperlink, table) on both `.odt` and `.docx`, verify no pipeline output leak (`$ret -isnot [array]`), verify XML structure inside both archives (including `<w:sectPr>`, `word/numbering.xml`, and `document.xml.rels` in `.docx`), and test `Convert-DocumentFormat` between `.odt`, `.docx`, and `.md`.

- [ ] **Step 2: Run `tests/test_converter.ps1` to verify RED failure**
  Run: `powershell -ExecutionPolicy Bypass -File tests/test_converter.ps1`
  Expected: FAIL (`doc_converter.ps1` not found).

- [ ] **Step 3: Implement `tests/create_sample.ps1` and `sidecars/doc_editor/doc_converter.ps1`**
  - Update `tests/create_sample.ps1` to support generating both valid `.odt` and valid `.docx` sample files based on file extension.
  - Implement `sidecars/doc_editor/doc_converter.ps1` with full ODT and DOCX readers, writers, blank container generators, dynamic style / relationship / numbering builders, and Markdown exporter.

- [ ] **Step 4: Run `tests/test_converter.ps1` to verify GREEN pass**
  Run: `powershell -ExecutionPolicy Bypass -File tests/test_converter.ps1`
  Expected: PASS.

---

### Task 2: Unified Sidecar HTTP Server (`server.ps1`)

**Files:**
- Create: `sidecars/doc_editor/server.ps1`
- Modify: `tests/test_server.ps1`

**Interfaces:**
- Consumes: `sidecars/doc_editor/doc_converter.ps1`
- Produces HTTP endpoints:
  - `GET /`, `/ui.html`, `/editor.css`, `/editor.js`
  - `GET /api/documents` -> `{ files: [ { name, path, format, size, lastModified } ] }` (scans for `*.odt` and `*.docx`)
  - `GET /api/document?path=<path>` -> `{ success, title, format, html, path }`
  - `POST /api/document` -> `{ path, html, createNew }` -> saves or creates `.odt` / `.docx`
  - `POST /api/convert` -> `{ sourcePath, targetPath, html, format }` -> saves converted `.docx`, `.odt`, or `.md` to `targetPath`
  - `GET /api/shutdown` -> local shutdown endpoint

- [ ] **Step 1: Write failing tests in `tests/test_server.ps1`**
  Update `tests/test_server.ps1` to test `/api/documents` (returning both `.odt` and `.docx` with `format` property), `GET /api/document` for both `.odt` and `.docx`, `POST /api/document` (saving edits and creating a new document), and `POST /api/convert` (converting `.odt` $\rightarrow$ `.docx`, `.docx` $\rightarrow$ `.odt`, and exporting `.md`).

- [ ] **Step 2: Run `tests/test_server.ps1` to verify RED failure**
  Run: `powershell -ExecutionPolicy Bypass -File tests/test_server.ps1`
  Expected: FAIL (`sidecars/doc_editor/server.ps1` not found).

- [ ] **Step 3: Implement `sidecars/doc_editor/server.ps1`**
  Implement the HTTP listener with `.odt`/`.docx` workspace discovery, document load/save/create, and `/api/convert` endpoint.

- [ ] **Step 4: Run `tests/test_server.ps1` to verify GREEN pass**
  Run: `powershell -ExecutionPolicy Bypass -File tests/test_server.ps1`
  Expected: PASS.

---

### Task 3: Full-Featured WYSIWYG UI, Menus, and Modals (`ui.html`, `editor.css`, `editor.js`)

**Files:**
- Create: `sidecars/doc_editor/ui.html`
- Create: `sidecars/doc_editor/editor.css`
- Create: `sidecars/doc_editor/editor.js`
- Modify: `tests/test_ui_assets.ps1`

**Interfaces:**
- Consumes: `/api/documents`, `/api/document`, `/api/convert`
- Produces: Complete interactive UI with all toolbar buttons, Save As / Export dropdown menu, New Document button, Insert Link / Table / Clear Formatting buttons, and custom modal dialog system.

- [ ] **Step 1: Write failing tests in `tests/test_ui_assets.ps1`**
  Update `tests/test_ui_assets.ps1` to verify `sidecars/doc_editor/ui.html`, `editor.css`, and `editor.js` exist and include all required controls:
  - Header controls: `#doc-title`, `#doc-format-badge`, `#doc-selector`, `#save-status`, `#btn-new`, `#btn-save`, `#btn-save-as`, `#save-as-menu`, `[data-export="docx"]`, `[data-export="odt"]`, `[data-export="md"]`
  - Toolbar controls: `#heading-select`, `#font-family-select`, `#font-size-select`, `#btn-bold`, `#btn-italic`, `#btn-underline`, `#btn-strike`, `#color-picker`, `#highlight-picker`, `#btn-align-left`, `#btn-align-center`, `#btn-align-right`, `#btn-align-justify`, `#btn-list-ul`, `#btn-list-ol`, `#btn-link`, `#btn-table`, `#btn-clear-format`, `#btn-undo`, `#btn-redo`
  - Modal elements: `#modal-overlay`, `#modal-title`, `#modal-body`, `#modal-confirm`, `#modal-cancel`
  - JS functions: `loadDocument`, `saveDocument`, `convertDocument`, `createNewDocument`, `openModal`, `insertLink`, `insertTable`, `updateToolbarState`.

- [ ] **Step 2: Run `tests/test_ui_assets.ps1` to verify RED failure**
  Run: `powershell -ExecutionPolicy Bypass -File tests/test_ui_assets.ps1`
  Expected: FAIL.

- [ ] **Step 3: Implement `ui.html`, `editor.css`, and `editor.js`**
  Build the complete UI, styling, dropdown menus, selection preservation for modals, link clickable tooltip/opening support, table insertion, active state tracking on cursor movement, and Save / Save Copy As / Export workflows.

- [ ] **Step 4: Run `tests/test_ui_assets.ps1` to verify GREEN pass**
  Run: `powershell -ExecutionPolicy Bypass -File tests/test_ui_assets.ps1`
  Expected: PASS.

---

### Task 4: Plugin Manifest, Sidecar Config, Skills, Rules & E2E Migration

**Files:**
- Modify: `plugin.json`
- Create: `sidecars/doc_editor/sidecar.json`
- Create: `skills/doc-editor/SKILL.md`
- Modify: `rules/AGENTS.md`
- Modify: `README.md`
- Modify: `tests/test_plugin_manifest.ps1`
- Modify: `tests/test_e2e.ps1`

- [ ] **Step 1: Write failing tests in `tests/test_plugin_manifest.ps1` and `tests/test_e2e.ps1`**
  Update `test_plugin_manifest.ps1` to assert plugin name `doc-editor`, sidecar `sidecars/doc_editor/sidecar.json`, skill `skills/doc-editor/SKILL.md`, and sidecar URI `sidecar://doc-editor/doc_editor/ui.html`. Update `test_e2e.ps1` to run end-to-end verification across both `.odt` and `.docx` plus cross-format conversions.

- [ ] **Step 2: Run `tests/test_plugin_manifest.ps1` to verify RED failure**
  Run: `powershell -ExecutionPolicy Bypass -File tests/test_plugin_manifest.ps1`
  Expected: FAIL.

- [ ] **Step 3: Update manifests, skills, rules, README, and remove legacy `odt_editor` paths**
  Update `plugin.json`, `sidecars/doc_editor/sidecar.json`, `skills/doc-editor/SKILL.md`, `rules/AGENTS.md`, `README.md`, and remove old `sidecars/odt_editor` and `skills/odt-editor` folders. Sync to `~/.gemini/config/plugins/doc-editor`, retire `~/.gemini/config/plugins/odt-editor`, rename the GitHub repo to `antigravity-doc-editor`, and push.

- [ ] **Step 4: Run full test suite to verify GREEN pass**
  Run: `powershell -ExecutionPolicy Bypass -File tests/test_e2e.ps1`
  Expected: PASS.
