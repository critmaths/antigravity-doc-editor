---
name: doc-editor
description: "Interactive visual document editor and viewer for Microsoft Word (.docx) and OpenDocument Text (.odt) documents in Antigravity with live disk saving, typography controls, and cross-format conversion."
---

# Antigravity Document Editor (`doc-editor`)

Antigravity Document Editor is a full-featured in-IDE document editing sidecar for Google Antigravity. It provides a visual WYSIWYG editing canvas, rich typography controls, table insertion, live saving back to disk, and instant cross-format conversion between **Word (`.docx`)**, **OpenDocument (`.odt`)**, and **Markdown (`.md`)**.

## Supported Document Formats
- **Microsoft Word OpenXML (`.docx`)**: Full round-trip editing preserving document body styles, external hyperlinks, bullet/numbered lists, and section page setups.
- **OASIS OpenDocument Text (`.odt`)**: Full round-trip editing preserving XML automatic styles, headings, quotes, tables, and hyperlinks.
- **Markdown (`.md`)**: Direct export of documents or editor content to clean Markdown.

## How to Open in Antigravity Sidecar

Whenever the user asks to preview, read, edit, or open a `.docx` or `.odt` document, provide a clickable sidecar link formatted as:

```markdown
[Open in Antigravity Document Editor](sidecar://doc-editor/doc_editor/ui.html?file=<URL_ENCODED_PATH>)
```

To open the blank editor canvas or choose from discovered documents:
```markdown
[Open Antigravity Document Editor](sidecar://doc-editor/doc_editor/ui.html)
```

## Features & Controls

1. **WYSIWYG Editing Canvas**:
   - Clean page layout with true-to-print margins.
   - Live content editable canvas with instant visual feedback.

2. **Typography & Styling Toolbar**:
   - **Paragraph Styles**: Normal Paragraph, Heading 1, Heading 2, Heading 3, Blockquote.
   - **Fonts & Sizes**: Calibri, Arial, Times New Roman, Georgia, Segoe UI, Courier New (9pt to 32pt).
   - **Inline Formatting**: Bold (`Ctrl+B`), Italic (`Ctrl+I`), Underline (`Ctrl+U`), Strikethrough (`<s>`).
   - **Coloring**: Native color picker for foreground text color and background highlight color.
   - **Alignment**: Align Left, Align Center, Align Right, Justify.
   - **Lists & Blocks**: Unordered bullet lists, ordered numbered lists, tables, and hyperlinks.
   - **History**: Undo (`Ctrl+Z`), Redo (`Ctrl+Y`), and Clear Formatting.

3. **Live Disk Saving**:
   - `Ctrl+S` or clicking the **Save** button immediately serializes the DOM back into the native OpenXML (`.docx`) or OpenDocument (`.odt`) package on disk without external dependencies.
   - Visual save status badge (`Saved`, `Unsaved changes`, `Saving...`).

4. **New Document & Export**:
   - **New**: Quickly scaffold a brand-new blank `.docx` or `.odt` file right from the IDE.
   - **Export**: Seamlessly convert open documents to `.docx`, `.odt`, or `.md` via the **Export** menu.

## CLI Converter Reference

If you only need to read and summarize a document in markdown or plain text via command line without opening the GUI sidecar:

```powershell
powershell -ExecutionPolicy Bypass -File "$HOME\.gemini\config\plugins\doc-editor\sidecars\doc_editor\doc_converter.ps1" -Read "path\to\document.docx" -Format markdown
```
