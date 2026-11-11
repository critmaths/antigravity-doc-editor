# Test for UI assets in doc_editor
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$SidecarDir = Join-Path $RepoDir "sidecars\doc_editor"
$UiFile = Join-Path $SidecarDir "ui.html"
$CssFile = Join-Path $SidecarDir "editor.css"
$JsFile = Join-Path $SidecarDir "editor.js"

Write-Host "Checking if UI files exist in $SidecarDir..."
if (-not (Test-Path $UiFile)) { Write-Error "FAIL: $UiFile does not exist" }
if (-not (Test-Path $CssFile)) { Write-Error "FAIL: $CssFile does not exist" }
if (-not (Test-Path $JsFile)) { Write-Error "FAIL: $JsFile does not exist" }

$uiContent = Get-Content $UiFile -Raw
$cssContent = Get-Content $CssFile -Raw
$jsContent = Get-Content $JsFile -Raw

# Check required UI elements in ui.html
$requiredIds = @(
    "doc-title", "doc-format-badge", "doc-selector", "save-status",
    "btn-open", "file-import-input", "folder-import-input", "btn-new", "btn-save", "btn-save-as", "save-as-menu",
    "heading-select", "font-family-select", "font-size-select",
    "btn-bold", "btn-italic", "btn-underline", "btn-strike",
    "color-picker", "highlight-picker",
    "btn-align-left", "btn-align-center", "btn-align-right", "btn-align-justify",
    "btn-list-ul", "btn-list-ol", "btn-link", "btn-table", "btn-clear-format",
    "btn-undo", "btn-redo",
    "editor-canvas",
    "modal-overlay", "modal-title", "modal-body", "modal-confirm", "modal-cancel"
)
foreach ($id in $requiredIds) {
    if ($uiContent -notmatch "id=['`"]$id['`"]") {
        Write-Error "FAIL: ui.html is missing element with id='$id'"
    }
}

# Check export format options in ui.html
$exportFormats = @("docx", "odt", "md")
foreach ($fmt in $exportFormats) {
    if ($uiContent -notmatch "data-export=['`"]$fmt['`"]") {
        Write-Error "FAIL: ui.html is missing data-export='$fmt'"
    }
}

# Check essential CSS selectors in editor.css
$requiredSelectors = @(
    "\.paper-sheet", "\.toolbar", "\.btn-tool", "#editor-canvas",
    "\.modal-overlay", "\.modal-dialog", "\.format-badge", "\.dropdown-menu"
)
foreach ($sel in $requiredSelectors) {
    if ($cssContent -notmatch $sel) {
        Write-Error "FAIL: editor.css is missing selector matching $sel"
    }
}

# Check essential JS functions in editor.js
$requiredJs = @(
    "loadDocument", "saveDocument", "convertDocument", "createNewDocument", "openDocumentModal",
    "openModal", "closeModal", "insertLink", "insertTable", "updateToolbarState"
)
foreach ($fn in $requiredJs) {
    if ($jsContent -notmatch $fn) {
        Write-Error "FAIL: editor.js is missing function or logic matching $fn"
    }
}

Write-Host "PASS: Task 3 UI asset tests passed!"
