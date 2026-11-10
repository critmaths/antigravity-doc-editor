# ODT Document Editing and Preview Rules

When working with OpenDocument Text (`.odt`) files in Antigravity:

1. **Always surface the sidecar editor link**: Whenever the user asks to preview, read, edit, or check an `.odt` file, include a clickable `sidecar://` pill link in your response:
   ```markdown
   [Open in ODT Document Editor](sidecar://odt-editor/odt_editor/ui.html?file=<URL_ENCODED_PATH>)
   ```

2. **Read documents via CLI**: If the user asks about the document's contents, use `powershell -ExecutionPolicy Bypass -File "$HOME\.gemini\config\plugins\odt-editor\sidecars\odt_editor\odt_converter.ps1" -Read "<PATH>" -Format markdown` to read and summarize it accurately.
