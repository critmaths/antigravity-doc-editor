# Test for Antigravity plugin manifest, sidecar configuration, skill, and rules
$ErrorActionPreference = "Stop"

$PluginDir = "C:\Users\critm\.gemini\config\plugins\odt-editor"
$PluginJson = Join-Path $PluginDir "plugin.json"
$SidecarJson = Join-Path $PluginDir "sidecars\odt_editor\sidecar.json"
$SkillMd = Join-Path $PluginDir "skills\odt-editor\SKILL.md"
$AgentsMd = Join-Path $PluginDir "rules\AGENTS.md"

Write-Host "Verifying plugin manifest files..."
if (-not (Test-Path $PluginJson)) { Write-Error "FAIL: $PluginJson missing" }
if (-not (Test-Path $SidecarJson)) { Write-Error "FAIL: $SidecarJson missing" }
if (-not (Test-Path $SkillMd)) { Write-Error "FAIL: $SkillMd missing" }
if (-not (Test-Path $AgentsMd)) { Write-Error "FAIL: $AgentsMd missing" }

# Validate plugin.json
$pJson = Get-Content $PluginJson -Raw | ConvertFrom-Json
if ($pJson.name -ne "odt-editor") { Write-Error "FAIL: plugin.json name is not 'odt-editor'" }
if (-not $pJson.description) { Write-Error "FAIL: plugin.json missing description" }

# Validate sidecar.json
$sJson = Get-Content $SidecarJson -Raw | ConvertFrom-Json
if ($sJson.command -ne "powershell") { Write-Error "FAIL: sidecar.json command is not powershell" }
if (-not $sJson.has_web_ui) { Write-Error "FAIL: sidecar.json has_web_ui is not true" }
if ($sJson.ui_config.views[0].entrypoint -ne "SIDECAR_UI_ENTRYPOINT_AUX_PANE") {
    Write-Error "FAIL: entrypoint is not SIDECAR_UI_ENTRYPOINT_AUX_PANE"
}
if ($sJson.ui_config.views[0].path -ne "/ui.html") {
    Write-Error "FAIL: view path is not /ui.html"
}

# Validate SKILL.md
$skillContent = Get-Content $SkillMd -Raw
if ($skillContent -notmatch "name:\s*odt-editor") { Write-Error "FAIL: SKILL.md missing name frontmatter" }
if ($skillContent -notmatch "sidecar://odt-editor/odt_editor/ui\.html") { Write-Error "FAIL: SKILL.md missing sidecar URL" }

Write-Host "PASS: Task 4 plugin manifest and configuration tests passed!"
