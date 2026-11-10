# Test for UI assets in odt_editor
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$SidecarDir = Join-Path $RepoDir "sidecars\odt_editor"
$UiFile = Join-Path $SidecarDir "ui.html"
$CssFile = Join-Path $SidecarDir "editor.css"
$JsFile = Join-Path $SidecarDir "editor.js"

Write-Host "Checking if UI files exist..."
if (-not (Test-Path $UiFile)) { Write-Error "FAIL: $UiFile does not exist" }
if (-not (Test-Path $CssFile)) { Write-Error "FAIL: $CssFile does not exist" }
if (-not (Test-Path $JsFile)) { Write-Error "FAIL: $JsFile does not exist" }

$uiContent = Get-Content $UiFile -Raw
$cssContent = Get-Content $CssFile -Raw
$jsContent = Get-Content $JsFile -Raw

# Check essential UI elements in ui.html
$requiredIds = @("doc-title", "font-family-select", "font-size-select", "heading-select", "btn-bold", "btn-italic", "btn-underline", "color-picker", "editor-canvas", "btn-save")
foreach ($id in $requiredIds) {
    if ($uiContent -notmatch "id=['`"]$id['`"]") {
        Write-Error "FAIL: ui.html is missing element with id='$id'"
    }
}

# Check essential CSS selectors in editor.css
$requiredSelectors = @("\.paper-sheet", "\.toolbar", "\.btn-tool", "#editor-canvas")
foreach ($sel in $requiredSelectors) {
    if ($cssContent -notmatch $sel) {
        Write-Error "FAIL: editor.css is missing selector matching $sel"
    }
}

# Check essential JS functions in editor.js
$requiredJs = @("loadDocument", "saveDocument", "formatText", "execCommand")
foreach ($fn in $requiredJs) {
    if ($jsContent -notmatch $fn) {
        Write-Error "FAIL: editor.js is missing function or logic matching $fn"
    }
}

Write-Host "PASS: Task 3 UI asset tests passed!"
