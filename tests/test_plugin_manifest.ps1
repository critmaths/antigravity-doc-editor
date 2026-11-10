# Test for Antigravity plugin manifest, sidecar configuration, skill, and rules (doc-editor)
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PluginDir = Split-Path -Parent $ScriptDir
$PluginJson = Join-Path $PluginDir "plugin.json"
$SidecarJson = Join-Path $PluginDir "sidecars\doc_editor\sidecar.json"
$SkillMd = Join-Path $PluginDir "skills\doc-editor\SKILL.md"
$AgentsMd = Join-Path $PluginDir "rules\AGENTS.md"

Write-Host "Verifying plugin manifest files for doc-editor..."
if (-not (Test-Path $PluginJson)) { Write-Error "FAIL: $PluginJson missing" }
if (-not (Test-Path $SidecarJson)) { Write-Error "FAIL: $SidecarJson missing" }
if (-not (Test-Path $SkillMd)) { Write-Error "FAIL: $SkillMd missing" }
if (-not (Test-Path $AgentsMd)) { Write-Error "FAIL: $AgentsMd missing" }

# Validate plugin.json
$pJson = Get-Content $PluginJson -Raw | ConvertFrom-Json
if ($pJson.name -ne "doc-editor") { Write-Error "FAIL: plugin.json name is not 'doc-editor'" }
if ($pJson.displayName -ne "Antigravity Document Editor") { Write-Error "FAIL: plugin.json displayName is not 'Antigravity Document Editor'" }
if (-not $pJson.description) { Write-Error "FAIL: plugin.json missing description" }

# Validate sidecar.json
$sJson = Get-Content $SidecarJson -Raw | ConvertFrom-Json
if ($sJson.name -ne "doc_editor") { Write-Error "FAIL: sidecar.json name is not doc_editor" }
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
if ($skillContent -notmatch "name:\s*doc-editor") { Write-Error "FAIL: SKILL.md missing name frontmatter" }
if ($skillContent -notmatch "sidecar://doc-editor/doc_editor/ui\.html") { Write-Error "FAIL: SKILL.md missing sidecar URL" }
if ($skillContent -notmatch "\.docx") { Write-Error "FAIL: SKILL.md missing .docx documentation" }
if ($skillContent -notmatch "\.odt") { Write-Error "FAIL: SKILL.md missing .odt documentation" }

# Validate AGENTS.md
$agentsContent = Get-Content $AgentsMd -Raw
if ($agentsContent -notmatch "sidecar://doc-editor/doc_editor/ui\.html") { Write-Error "FAIL: AGENTS.md missing sidecar URL" }
if ($agentsContent -notmatch "doc_converter\.ps1") { Write-Error "FAIL: AGENTS.md missing doc_converter.ps1" }
if ($agentsContent -notmatch "\.docx") { Write-Error "FAIL: AGENTS.md missing .docx mention" }

Write-Host "PASS: Task 4 plugin manifest and configuration tests passed!"
