# End-to-End verification script for Antigravity Document Editor (doc-editor)
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PluginDir = Split-Path -Parent $ScriptDir
$ServerScript = Join-Path $PluginDir "sidecars\doc_editor\server.ps1"
$ConverterScript = Join-Path $PluginDir "sidecars\doc_editor\doc_converter.ps1"
$CreateSampleScript = Join-Path $ScriptDir "create_sample.ps1"
$TargetOdt = Join-Path $ScriptDir "sample.odt"
$TargetDocx = Join-Path $ScriptDir "sample.docx"
$BackupOdt = Join-Path $ScriptDir "sample.odt.bak"
$BackupDocx = Join-Path $ScriptDir "sample.docx.bak"
$CreatedNewDocx = Join-Path $ScriptDir "e2e_created.docx"
$ConvertedDocx = Join-Path $ScriptDir "e2e_converted.docx"
$ConvertedMd = Join-Path $ScriptDir "e2e_exported.md"

Write-Host "=== Step 1: Creating fresh generic sample ODT and DOCX ==="
& $CreateSampleScript -OutputPath $TargetOdt
& $CreateSampleScript -OutputPath $TargetDocx
Copy-Item -Path $TargetOdt -Destination $BackupOdt -Force
Copy-Item -Path $TargetDocx -Destination $BackupDocx -Force

try {
    Write-Host "=== Step 2: Testing CLI converter on sample documents ==="
    $cliOutputOdt = & $ConverterScript -Read $TargetOdt -Format markdown
    if ($cliOutputOdt -notmatch "Sample Document" -or $cliOutputOdt -notmatch "Overview") {
        Write-Error "FAIL: CLI converter did not extract expected sections from sample ODT"
    }

    $cliOutputDocx = & $ConverterScript -Read $TargetDocx -Format markdown
    if ($cliOutputDocx -notmatch "Sample Document" -or $cliOutputDocx -notmatch "Overview") {
        Write-Error "FAIL: CLI converter did not extract expected sections from sample DOCX"
    }
    Write-Host "CLI extraction successful for both ODT and DOCX."

    Write-Host "=== Step 3: Starting Sidecar Server ==="
    $TestPort = 58930
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
    Write-Host "Found $($docs.files.Count) document files in search paths"

    Write-Host "=== Step 4: Loading documents via HTTP API ==="
    $encodedOdt = [System.Net.WebUtility]::UrlEncode($TargetOdt)
    $docJsonOdt = $wc.DownloadString("http://127.0.0.1:$TestPort/api/document?path=$encodedOdt")
    $docOdt = $docJsonOdt | ConvertFrom-Json
    if (-not $docOdt.html -or $docOdt.html -notmatch "Sample Document") {
        Write-Error "FAIL: HTTP API did not return valid ODT document HTML"
    }

    $encodedDocx = [System.Net.WebUtility]::UrlEncode($TargetDocx)
    $docJsonDocx = $wc.DownloadString("http://127.0.0.1:$TestPort/api/document?path=$encodedDocx")
    $docDocx = $docJsonDocx | ConvertFrom-Json
    if (-not $docDocx.html -or $docDocx.html -notmatch "Sample Document") {
        Write-Error "FAIL: HTTP API did not return valid DOCX document HTML"
    }
    Write-Host "Both ODT and DOCX loaded successfully via HTTP API."

    Write-Host "=== Step 5: Editing and Saving documents via HTTP API ==="
    $editedHtmlOdt = "<h1>Edited ODT via E2E Test</h1><p>Verified live save to disk with formatting.</p>"
    $saveBodyOdt = @{
        path = $TargetOdt
        html = $editedHtmlOdt
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $saveRespOdt = $wc.UploadString("http://127.0.0.1:$TestPort/api/document", "POST", $saveBodyOdt) | ConvertFrom-Json
    if (-not $saveRespOdt.success) {
        Write-Error "FAIL: POST /api/document failed for ODT"
    }

    $editedHtmlDocx = "<h1>Edited DOCX via E2E Test</h1><p>Verified live save to disk with formatting.</p>"
    $saveBodyDocx = @{
        path = $TargetDocx
        html = $editedHtmlDocx
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $saveRespDocx = $wc.UploadString("http://127.0.0.1:$TestPort/api/document", "POST", $saveBodyDocx) | ConvertFrom-Json
    if (-not $saveRespDocx.success) {
        Write-Error "FAIL: POST /api/document failed for DOCX"
    }

    # Verify disk files reflect changes
    $recheckOdt = & $ConverterScript -Read $TargetOdt -Format markdown
    if ($recheckOdt -notmatch "Edited ODT via E2E Test") {
        Write-Error "FAIL: ODT disk file was not updated properly by HTTP API"
    }

    $recheckDocx = & $ConverterScript -Read $TargetDocx -Format markdown
    if ($recheckDocx -notmatch "Edited DOCX via E2E Test") {
        Write-Error "FAIL: DOCX disk file was not updated properly by HTTP API"
    }
    Write-Host "Verified persistent disk updates for both formats."

    Write-Host "=== Step 6: Testing New Document & Conversion via HTTP API ==="
    if (Test-Path $CreatedNewDocx) { Remove-Item -Path $CreatedNewDocx -Force }
    $createBody = @{
        createNew = $true
        path = $CreatedNewDocx
        title = "E2E Created Doc"
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $createResp = $wc.UploadString("http://127.0.0.1:$TestPort/api/document", "POST", $createBody) | ConvertFrom-Json
    if (-not $createResp.success -or -not (Test-Path $CreatedNewDocx)) {
        Write-Error "FAIL: New document creation failed in E2E test"
    }

    if (Test-Path $ConvertedDocx) { Remove-Item -Path $ConvertedDocx -Force }
    $convBody = @{
        sourcePath = $TargetOdt
        targetPath = $ConvertedDocx
        format = "docx"
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $convResp = $wc.UploadString("http://127.0.0.1:$TestPort/api/convert", "POST", $convBody) | ConvertFrom-Json
    if (-not $convResp.success -or -not (Test-Path $ConvertedDocx)) {
        Write-Error "FAIL: ODT -> DOCX conversion failed in E2E test"
    }

    if (Test-Path $ConvertedMd) { Remove-Item -Path $ConvertedMd -Force }
    $mdBody = @{
        html = "<h1>E2E Markdown Export</h1><p>Exporting to clean markdown text.</p>"
        targetPath = $ConvertedMd
        format = "md"
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $mdResp = $wc.UploadString("http://127.0.0.1:$TestPort/api/convert", "POST", $mdBody) | ConvertFrom-Json
    if (-not $mdResp.success -or -not (Test-Path $ConvertedMd)) {
        Write-Error "FAIL: Markdown export failed in E2E test"
    }

    Write-Host "PASS: Antigravity Document Editor end-to-end verification succeeded 100%!"

} finally {
    try {
        $wc = New-Object System.Net.WebClient
        $wc.DownloadString("http://127.0.0.1:$TestPort/api/shutdown") | Out-Null
    } catch {}
    if ($job) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -Path $BackupOdt, $BackupDocx, $CreatedNewDocx, $ConvertedDocx, $ConvertedMd -Force -ErrorAction SilentlyContinue
}
