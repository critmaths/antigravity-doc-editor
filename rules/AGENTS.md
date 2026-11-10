# Antigravity Document Editor Operating Rules

When working with Microsoft Word (`.docx`) and OpenDocument Text (`.odt`) files in Antigravity:

1. **Always surface the sidecar editor link**: Whenever the user asks to preview, read, edit, or check a `.docx` or `.odt` document, include a clickable `sidecar://` link in your response:
   ```markdown
   [Open in Antigravity Document Editor](sidecar://doc-editor/doc_editor/ui.html?file=<URL_ENCODED_PATH>)
   ```
   Or without a preselected file:
   ```markdown
   [Open Antigravity Document Editor](sidecar://doc-editor/doc_editor/ui.html)
   ```

2. **Read documents via CLI**: If the user asks for a summary or questions about the document's contents without needing the visual UI, use the unified converter CLI to read and summarize it accurately:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "$HOME\.gemini\config\plugins\doc-editor\sidecars\doc_editor\doc_converter.ps1" -Read "<PATH>" -Format markdown
   ```
