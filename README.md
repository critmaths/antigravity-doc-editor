# Antigravity ODT Viewer & Editor Plugin

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Antigravity%202.0%20%7C%20Windows-brightgreen.svg)]()
[![Zero Dependencies](https://img.shields.io/badge/dependencies-zero%20external-orange.svg)]()

> A native **Antigravity** plugin that brings in-IDE previewing and WYSIWYG editing for OpenDocument Text (`.odt`) files directly into Antigravity's auxiliary side panel, with instant live disk persistence.

---

## 💡 The Problem

In Antigravity 2.0, clicking an `.odt` file in your workspace tree displays:

> *"Preview is not available for this file type. Download document.odt"*

Developers and writers are forced to launch an external office suite like LibreOffice Writer or Microsoft Word just to read an outline or tweak project notes.

## ✨ The Solution

**Antigravity ODT Editor** runs directly inside Antigravity's auxiliary side pane. It gives you a clean, paper-style document canvas and rich word-processing controls side-by-side with your code and conversation tabs.

- **Zero External Dependencies**: Built to run out of the box on Windows using PowerShell and native .NET `System.Net.HttpListener` & `System.IO.Compression.ZipFile`. No Node.js, Python, or npm installations required.
- **Bi-Directional ODF Preservation**: Parses `.odt` ZIP packages, converts formatting into editable HTML, and serializes user edits back into valid OASIS OpenDocument XML with automatic styles. Non-body entries (`styles.xml`, `meta.xml`, `manifest.xml`, `Pictures/`) remain untouched.
- **Rich Text Formatting Toolbar**:
  - **Fonts**: System sans, serif, monospace, plus standard document typography (Inter, Calibri, Arial, Times New Roman, Georgia).
  - **Sizes**: 10pt, 12pt, 14pt, 18pt, 24pt, 32pt.
  - **Styles**: Bold, Italic, Underline, Strikethrough.
  - **Structure**: Normal Text, Headings (H1–H3), Blockquotes.
  - **Colors**: Foreground text color and background highlight color pickers.
  - **Alignment**: Left, Center, Right, Justify.
  - **Lists**: Bulleted and Numbered lists.
  - **Save**: Fast **Ctrl+S** keyboard shortcut and Save button with dirty-state indicator.
- **Workspace Switcher**: Automatically discovers and lists all `.odt` files in your workspace for one-click switching.
- **Autonomous Agent Integration**: Bundled skill and context rules that allow the AI agent to surface preview links and programmatically read/summarize documents via CLI.

---

## 🚀 Installation

Clone or copy this repository into your Antigravity plugins directory:

```powershell
git clone https://github.com/critmaths/antigravity-odt-editor.git "$HOME\.gemini\config\plugins\odt-editor"
```

The plugin will automatically be discovered by Antigravity and ready for use immediately without requiring an IDE restart.

---

## 📖 Usage

### In Antigravity Chat
When working with any `.odt` file, click the sidecar pill link:

```markdown
[Open in ODT Document Editor](sidecar://odt-editor/odt_editor/ui.html?file=<URL_ENCODED_PATH>)
```

### From the Customizations Menu
Open `Antigravity Settings` > `Customizations` to view or toggle the plugin.

### CLI Mode (Scripting & Agent Inspection)
You can also inspect or extract `.odt` text directly from PowerShell:

```powershell
# Extract formatted Markdown
powershell -ExecutionPolicy Bypass -File "sidecars/odt_editor/odt_converter.ps1" -Read "path/to/doc.odt" -Format markdown

# Extract clean plain text
powershell -ExecutionPolicy Bypass -File "sidecars/odt_editor/odt_converter.ps1" -Read "path/to/doc.odt" -Format text
```

---

## 🧪 Automated Test Suite

Run the full end-to-end test suite from PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File "tests/test_e2e.ps1"
```

All 5 core test suites verify:
1. `test_converter.ps1`: Parsing, HTML generation, and round-trip ODF XML serialization.
2. `test_server.ps1`: HTTP endpoints, CORS headers, document retrieval, and REST save.
3. `test_ui_assets.ps1`: UI element IDs, CSS themes, and client WYSIWYG commands.
4. `test_plugin_manifest.ps1`: Antigravity manifest schema, sidecar definition, and skill metadata.
5. `test_e2e.ps1`: Full pipeline modification and disk verification on sample documents.

---

## 📜 License

MIT License © 2026 [Scott Forest (@critmaths)](https://github.com/critmaths).
