# Antigravity Document Editor (`doc-editor`)

[![Release](https://img.shields.io/badge/version-2.3.0-blue.svg)](https://github.com/critmaths/antigravity-doc-editor/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Google%20Antigravity%20%7C%20Windows-brightgreen.svg)]()
[![Zero Dependencies](https://img.shields.io/badge/Dependencies-Zero%20External%20(Pure%20.NET)-orange.svg)]()
[![Supported Formats](https://img.shields.io/badge/Formats-.docx%20%7C%20.odt%20%7C%20.md%20%7C%20.html%20%7C%20.txt-purple.svg)]()

> An embedded **WYSIWYG document editor and viewer** for **Google Antigravity**. Read, edit, format, and convert **Microsoft Word (`.docx`)** and **OpenDocument Text (`.odt`)** files directly in your IDE auxiliary pane with live disk synchronization and Markdown export.

---

## 💡 Why This Exists

When an AI pair programmer or workflow generates `.docx` or `.odt` technical specs, reports, or documentation, clicking them inside Antigravity displays:

> *"Preview is not available for this file type."*

Having to leave your editor, launch heavy desktop office software (Microsoft Word or LibreOffice Writer), make tweaks, and remember to save breaks development flow.

**Antigravity Document Editor** embeds a visual, true-to-print editing canvas right inside your IDE side panel—running side-by-side with your code, terminal, and AI pair-programmer.

---

## ✨ Features

- 📝 **Native Word & OpenDocument Fidelity**:
  - **Microsoft Word OpenXML (`.docx`)**: Preserves `<w:sectPr>` page geometry, ensures `word/styles.xml` and `word/numbering.xml` integrity, registers hyperlinks in `word/_rels/document.xml.rels`, and formats headings, tables, and lists.
  - **OASIS OpenDocument Text (`.odt`)**: Automatically generates styles for paragraph alignment, colors, font sizes, blockquotes, and tables without corrupting meta or manifest entries.
- 📂 **Quick Open & File Import**:
  - Open local documents via the OS file dialog (<kbd>Ctrl</kbd> + <kbd>O</kbd>) or paste any disk path (with automatic Windows quotation mark sanitization).
  - Fast, non-blocking binary file import via native data streaming.
- 💾 **Interactive Directory Browser on Save As / Export**:
  - One-click quick location pills: **Documents**, **Desktop**, **Downloads**, **Scratch Workspace**, and all system drives (`C:\`, `D:\`).
  - Visual subfolder explorer with explicit **Select** and **Open ➔** navigation buttons and real-time destination preview.
- ⤓ **Multi-Format Export**:
  - Export any open document to **Word (`.docx`)**, **OpenDocument (`.odt`)**, **Markdown (`.md`)**, **HTML Webpage (`.html`)**, or **Plain Text (`.txt`)**.
- 🛠️ **Full WYSIWYG Editing Suite**:
  - **Headings & Hierarchy**: Normal Paragraph, Heading 1, Heading 2, Heading 3, Blockquotes.
  - **Typography**: Calibri, Arial, Times New Roman, Georgia, Segoe UI, Courier New (9pt to 32pt).
  - **Formatting**: Bold, Italic, Underline, Strikethrough, Text Color, and Highlight Markers.
  - **Structure**: Bulleted lists, numbered lists, tables, and hyperlinks.
  - **Undo/Redo**: Full state history with <kbd>Ctrl</kbd> + <kbd>Z</kbd> and <kbd>Ctrl</kbd> + <kbd>Y</kbd>.
- ⚡ **Zero Bloat, Zero Dependencies**:
  - Powered by native Windows PowerShell and built-in .NET classes (`System.Net.HttpListener`, `System.Xml`, `System.IO.Compression.ZipFile`).
  - **No Node.js, npm, Python, or Docker required.**
- 🤖 **Autonomous Agent Integration**:
  - Includes an Antigravity plugin manifest, agent skill (`SKILL.md`), and system rules (`AGENTS.md`) enabling the AI agent to read/summarize documents via CLI and surface one-click preview links in chat.

---

## ⚡ Quickstart & Installation

### Step 1: Install Plugin (1-Second Command)

In Windows PowerShell, run:

```powershell
git clone https://github.com/critmaths/antigravity-doc-editor.git "$HOME\.gemini\config\plugins\doc-editor"
```

Antigravity automatically discovers and loads the plugin immediately without requiring an IDE restart.

### Step 2: Open the Editor

Click the link below inside Antigravity to launch the editor:

👉 `sidecar://doc-editor/doc_editor/ui.html`

Or open a specific file directly:

👉 `sidecar://doc-editor/doc_editor/ui.html?file=C:\path\to\document.docx`

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
| :--- | :--- |
| <kbd>Ctrl</kbd> + <kbd>S</kbd> | Save document live to disk |
| <kbd>Ctrl</kbd> + <kbd>O</kbd> | Open / Import document modal |
| <kbd>Ctrl</kbd> + <kbd>Z</kbd> | Undo last edit |
| <kbd>Ctrl</kbd> + <kbd>Y</kbd> | Redo edit |
| <kbd>Ctrl</kbd> + <kbd>B</kbd> | Bold text |
| <kbd>Ctrl</kbd> + <kbd>I</kbd> | Italic text |
| <kbd>Ctrl</kbd> + <kbd>U</kbd> | Underline text |

---

## 💻 CLI Reader & Formatter

You can also use the bundled converter from the command line or scripts to read documents as Markdown, HTML, or plain text:

```powershell
powershell -ExecutionPolicy Bypass -File "$HOME\.gemini\config\plugins\doc-editor\sidecars\doc_editor\doc_converter.ps1" -Read "path\to\document.docx" -Format markdown
```

Supported CLI formats:
- `-Format markdown` (outputs clean Markdown text)
- `-Format html` (outputs rendered HTML)
- `-Format text` (outputs stripped plain text)

---

## 🧪 Automated Test Suite

The codebase includes 100% automated test coverage across parser fidelity, server endpoints, and end-to-end integration:

```powershell
# Run the complete test suite
powershell -ExecutionPolicy Bypass -File tests/test_converter.ps1
powershell -ExecutionPolicy Bypass -File tests/test_server.ps1
powershell -ExecutionPolicy Bypass -File tests/test_ui_assets.ps1
powershell -ExecutionPolicy Bypass -File tests/test_plugin_manifest.ps1
powershell -ExecutionPolicy Bypass -File tests/test_e2e.ps1
```

---

## 📁 Repository Structure

```
antigravity-doc-editor/
├── plugin.json                     # Antigravity plugin manifest
├── README.md                       # Documentation & quickstart
├── LICENSE                         # MIT License
├── rules/
│   └── AGENTS.md                   # Agent system rules for document handling
├── skills/
│   └── doc-editor/
│       └── SKILL.md                # Antigravity skill instructions
├── sidecars/
│   └── doc_editor/
│       ├── sidecar.json            # Antigravity sidecar configuration
│       ├── server.ps1              # Lightweight HTTP backend (port binding & APIs)
│       ├── doc_converter.ps1       # Pure .NET DOCX/ODT parser and exporter
│       ├── ui.html                 # WYSIWYG editor UI layout
│       ├── editor.css              # Dark-mode IDE styling & paper canvas
│       └── editor.js               # Client controller, toolbar, & disk sync
└── tests/                          # 100% automated test suites
    ├── test_converter.ps1
    ├── test_server.ps1
    ├── test_ui_assets.ps1
    ├── test_plugin_manifest.ps1
    └── test_e2e.ps1
```

---

## 📄 License

Released under the [MIT License](LICENSE). Contributions, bug reports, and pull requests are welcome!
