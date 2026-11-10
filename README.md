# Antigravity Document Editor (`doc-editor`)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Antigravity%202.0%20%7C%20Windows-brightgreen.svg)]()
[![Zero Dependencies](https://img.shields.io/badge/dependencies-zero%20external-orange.svg)]()

> A native **Antigravity** plugin that brings in-IDE visual previewing, WYSIWYG editing, live disk saving, and cross-format conversion for **Microsoft Word (`.docx`)** and **OpenDocument Text (`.odt`)** documents, plus **Markdown (`.md`)** export.

---

## 💡 The Problem

In Antigravity 2.0, clicking a `.docx` or `.odt` file in your workspace tree displays:

> *"Preview is not available for this file type."*

Developers, researchers, and technical writers are forced to context-switch out of Antigravity to launch external desktop office suites like Microsoft Word or LibreOffice Writer just to read a document, tweak documentation notes, or edit specifications.

## ✨ The Solution

**Antigravity Document Editor** runs directly inside Antigravity's auxiliary side panel. It delivers a clean, true-to-print paper canvas and rich typography controls side-by-side with your code, terminal, and AI pair-programmer.

- **Zero External Dependencies**: Runs out of the box on Windows using PowerShell and native .NET `System.Net.HttpListener`, `System.Xml.XmlDocument`, & `System.IO.Compression.ZipFile`. No Node.js, Python, or npm installations required.
- **Bi-Directional Format Fidelity**:
  - **Microsoft Word OpenXML (`.docx`)**: Preserves `<w:sectPr>` page geometry, ensures `word/numbering.xml` and `word/styles.xml` integrity, registers external links in `word/_rels/document.xml.rels`, and formats headings, paragraphs, tables, lists, and inline styles.
  - **OASIS OpenDocument Text (`.odt`)**: Generates automatic XML styles for paragraph alignments, colors, font sizes, headings, quotes, tables, and hyperlinks while preserving non-body entries.
- **Cross-Format Conversion & Export**:
  - Convert `.odt` $\leftrightarrow$ `.docx` instantly with one click.
  - Export any document directly to clean **Markdown (`.md`)**.
  - Create brand new `.docx` or `.odt` documents directly from the IDE.
- **Full WYSIWYG Toolbar**:
  - **Styles & Headings**: Normal Paragraph, Heading 1, Heading 2, Heading 3, Blockquote.
  - **Typography**: Calibri, Arial, Times New Roman, Georgia, Segoe UI, Courier New (9pt to 32pt).
  - **Inline Formats**: Bold (`Ctrl+B`), Italic (`Ctrl+I`), Underline (`Ctrl+U`), Strikethrough (`<s>`).
  - **Colors**: Foreground text color picker and background highlight marker.
  - **Alignment**: Align Left, Align Center, Align Right, Justify.
  - **Lists & Blocks**: Bulleted lists, numbered lists, tables, and hyperlinks.
  - **History**: Undo (`Ctrl+Z`), Redo (`Ctrl+Y`), and Clear Formatting.
  - **Save**: Fast **Ctrl+S** keyboard shortcut and Save button with live disk persistence.
- **Document Switcher**: Automatically discovers and lists all `.docx` and `.odt` files in your workspace and folders.
- **Custom In-Page Modals**: Native HTML modal dialogs for link insertion, table creation, document naming, and export targets (no blocked browser popups).
- **Autonomous Agent Integration**: Bundled skill and context rules that allow the AI agent to surface preview links and programmatically read/summarize documents via CLI.

---

## 🚀 Installation

Clone or copy this repository into your Antigravity plugins directory:

```powershell
git clone https://github.com/critmaths/antigravity-doc-editor.git "$HOME\.gemini\config\plugins\doc-editor"
```

The plugin will automatically be discovered by Antigravity and ready for use immediately without requiring an IDE restart.

---

## 🖥️ Usage

### 1. Visual In-IDE Editing
Whenever you ask your Antigravity AI pair-programmer to open, preview, or edit a document, it surfaces a clickable link:

```markdown
[Open in Antigravity Document Editor](sidecar://doc-editor/doc_editor/ui.html?file=<URL_ENCODED_PATH>)
```

Clicking the link opens the document in the auxiliary side panel with live bi-directional disk saving.

### 2. CLI Reader & Markdown Converter
To read and inspect document contents from the command line:

```powershell
powershell -ExecutionPolicy Bypass -File "sidecars\doc_editor\doc_converter.ps1" -Read "path\to\document.docx" -Format markdown
```

Supported output formats: `markdown`, `html`, `text`.

---

## 🧪 Testing

The repository includes a comprehensive 100% automated test suite:

```powershell
# Run all unit and integration tests
powershell -ExecutionPolicy Bypass -File tests/test_converter.ps1
powershell -ExecutionPolicy Bypass -File tests/test_server.ps1
powershell -ExecutionPolicy Bypass -File tests/test_ui_assets.ps1
powershell -ExecutionPolicy Bypass -File tests/test_plugin_manifest.ps1
powershell -ExecutionPolicy Bypass -File tests/test_e2e.ps1
```

---

## 📄 License

MIT License. See [LICENSE](LICENSE) for details.
