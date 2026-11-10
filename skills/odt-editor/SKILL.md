---
name: odt-editor
description: Preview, format, and edit OpenDocument Text (.odt) documents in Antigravity. Use when the user asks to preview, read, edit, or open an .odt file, or when working with LibreOffice / OpenOffice documents.
---

# ODT Document Editor & Preview

This skill allows Antigravity to preview and edit OpenDocument Text (`.odt`) files directly within the IDE auxiliary pane.

## 1. Opening the In-IDE Editor

When the user asks to **open**, **preview**, **edit**, or **view** an `.odt` file, always provide the clickable sidecar link so they can open the editor in the auxiliary pane:

```markdown
[ODT Document Editor](sidecar://odt-editor/odt_editor/ui.html?file=<URL_ENCODED_PATH>)
```

Example for a document in the workspace:
```markdown
[Open in ODT Editor](sidecar://odt-editor/odt_editor/ui.html?file=C%3A%2Fpath%2Fto%2Fsample.odt)
```

## 2. Inspecting Document Content in CLI

When you need to read the contents of an `.odt` file to answer questions, summarize, or draft text, execute the converter CLI directly:

```powershell
powershell -ExecutionPolicy Bypass -File "$HOME\.gemini\config\plugins\odt-editor\sidecars\odt_editor\odt_converter.ps1" -Read "<PATH_TO_ODT>" -Format markdown
```

Formats supported:
- `markdown` (default): Extracts headings, paragraphs, lists, and formatting.
- `text`: Strips all tags to plain text.
- `html`: Raw document HTML representation.

## 3. Formatting Features in Editor

The in-IDE editor supports:
- Font families (Inter, Calibri, Arial, Times New Roman, Georgia, Monospace)
- Font sizes (10pt, 12pt, 14pt, 18pt, 24pt, 32pt)
- Heading levels (Heading 1, 2, 3, Normal Text, Quotes)
- Text styles (Bold, Italic, Underline, Strikethrough)
- Colors (Text color, Background highlight)
- Alignment (Left, Center, Right, Justify)
- Lists (Bulleted, Numbered)
- Direct Ctrl+S saving to the original `.odt` file on disk.
