# End-to-End verification script for ODT Editor plugin
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PluginDir = Split-Path -Parent $ScriptDir
$ServerScript = Join-Path $PluginDir "sidecars\odt_editor\server.ps1"
$ConverterScript = Join-Path $PluginDir "sidecars\odt_editor\odt_converter.ps1"
$CreateSampleScript = Join-Path $ScriptDir "create_sample.ps1"
$TargetOdt = Join-Path $ScriptDir "sample.odt"
$BackupOdt = Join-Path $ScriptDir "sample.odt.bak"

Write-Host "=== Step 1: Creating fresh generic sample ODT ==="
& $CreateSampleScript -OutputPath $TargetOdt
Copy-Item -Path $TargetOdt -Destination $BackupOdt -Force

try {
    Write-Host "=== Step 2: Testing CLI converter on sample document ==="
    $cliOutput = & $ConverterScript -Read $TargetOdt -Format markdown
    if ($cliOutput -notmatch "Sample Document" -or $cliOutput -notmatch "Overview") {
        Write-Error "FAIL: CLI converter did not extract expected sections from sample ODT"
    }
    Write-Host "CLI extraction successful. Found expected sample sections."

    Write-Host "=== Step 3: Starting Sidecar Server ==="
    $TestPort = 58929
    $job = Start-Job -ScriptBlock {
        param($s, $p)
        & $s -Port $p
    } -ArgumentList $ServerScript, $TestPort

    Start-Sleep -Seconds 2

    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("Accept", "application/json")

    # Verify server listening
    $docsJson = $wc.DownloadString("http://127.0.0.1:$TestPort/api/documents")
    $docs = $docsJson | ConvertFrom-Json
    Write-Host "Found $($docs.files.Count) ODT files in search paths"

    Write-Host "=== Step 4: Loading document via HTTP API ==="
    $encodedPath = [System.Net.WebUtility]::UrlEncode($TargetOdt)
    $docJson = $wc.DownloadString("http://127.0.0.1:$TestPort/api/document?path=$encodedPath")
    $doc = $docJson | ConvertFrom-Json
    if (-not $doc.html -or $doc.html -notmatch "Sample Document") {
        Write-Error "FAIL: HTTP API did not return document HTML"
    }
    Write-Host "Document loaded via HTTP API. Length: $($doc.html.Length) chars"

    Write-Host "=== Step 5: Editing and Saving document via HTTP API ==="
    $originalHtml = $doc.html
    $updatedHtml = $originalHtml + "<p><b>Updated by Antigravity ODT Editor Test on " + (Get-Date).ToString("g") + "</b></p>"
    
    $postBody = @{
        path = $TargetOdt
        html = $updatedHtml
    } | ConvertTo-Json

    $wc.Headers.Add("Content-Type", "application/json")
    $saveResult = $wc.UploadString("http://127.0.0.1:$TestPort/api/document", "POST", $postBody)
    $saveResp = $saveResult | ConvertFrom-Json
    if (-not $saveResp.success) {
        Write-Error "FAIL: Document save failed via HTTP API"
    }
    Write-Host "Document saved successfully."

    Write-Host "=== Step 6: Verifying saved modifications on disk ==="
    $verifyOutput = & $ConverterScript -Read $TargetOdt -Format markdown
    if ($verifyOutput -notmatch "Updated by Antigravity ODT Editor Test") {
        Write-Error "FAIL: Saved modifications not found in verified document"
    }
    Write-Host "Saved modification verified on disk!"

    Write-Host "=== ALL END-TO-END TESTS PASSED SUCCESSFULLY! ==="
} finally {
    try {
        $wc = New-Object System.Net.WebClient
        $wc.DownloadString("http://127.0.0.1:$TestPort/api/shutdown") | Out-Null
    } catch {}
    if ($job) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -Path $BackupOdt, $TargetOdt -Force -ErrorAction SilentlyContinue
}
