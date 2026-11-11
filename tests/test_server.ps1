# Test for doc_editor server.ps1 (ODT + DOCX + Blank creation + Cross-conversion API)
$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$ServerScript = Join-Path $RepoDir "sidecars\doc_editor\server.ps1"
$CreateSampleScript = Join-Path $ScriptDir "create_sample.ps1"
$SampleOdt = Join-Path $ScriptDir "sample.odt"
$SampleDocx = Join-Path $ScriptDir "sample.docx"
$TempOdt = Join-Path $ScriptDir "server_test_temp.odt"
$TempDocx = Join-Path $ScriptDir "server_test_temp.docx"
$CreatedNewDocx = Join-Path $ScriptDir "server_created_new.docx"
$ConvertedDocx = Join-Path $ScriptDir "server_converted.docx"
$ConvertedMd = Join-Path $ScriptDir "server_converted.md"

& $CreateSampleScript -OutputPath $SampleOdt
& $CreateSampleScript -OutputPath $SampleDocx
Copy-Item -Path $SampleOdt -Destination $TempOdt -Force
Copy-Item -Path $SampleDocx -Destination $TempDocx -Force

Write-Host "Checking if server script exists: $ServerScript"
if (-not (Test-Path $ServerScript)) {
    Write-Error "FAIL: $ServerScript does not exist"
}

$TestPort = 58926
$job = Start-Job -ScriptBlock {
    param($s, $p)
    & $s -Port $p
} -ArgumentList $ServerScript, $TestPort

try {
    # Wait for server to start listening
    $started = $false
    for ($i = 0; $i -lt 15; $i++) {
        Start-Sleep -Milliseconds 400
        try {
            $wc = New-Object System.Net.WebClient
            $wc.Headers.Add("Accept", "application/json")
            $respStr = $wc.DownloadString("http://127.0.0.1:$TestPort/api/documents")
            if ($respStr) {
                $started = $true
                break
            }
        } catch {}
    }

    if (-not $started) {
        Write-Error "FAIL: Server did not respond within timeout"
    }

    Write-Host "Server started successfully on port $TestPort"

    # Test 1: GET /api/documents
    $wc = New-Object System.Net.WebClient
    $docsJson = $wc.DownloadString("http://127.0.0.1:$TestPort/api/documents")
    $docsResp = $docsJson | ConvertFrom-Json
    if (-not $docsResp.files -or $docsResp.files.Count -eq 0) {
        Write-Error "FAIL: /api/documents returned no files"
    }
    $formatsFound = @($docsResp.files | ForEach-Object { $_.format })
    if ($formatsFound -notcontains "odt" -or $formatsFound -notcontains "docx") {
        Write-Error "FAIL: /api/documents did not discover both odt and docx files"
    }
    Write-Host "GET /api/documents returned $($docsResp.files.Count) files (found both odt and docx)."

    # Test 2: GET /api/document?path=... (both ODT and DOCX)
    $encodedOdt = [System.Net.WebUtility]::UrlEncode($TempOdt)
    $docJson = $wc.DownloadString("http://127.0.0.1:$TestPort/api/document?path=$encodedOdt")
    $odtResp = $docJson | ConvertFrom-Json
    if (-not $odtResp.html -or $odtResp.format -ne "odt") {
        Write-Error "FAIL: GET /api/document failed on ODT"
    }

    $encodedDocx = [System.Net.WebUtility]::UrlEncode($TempDocx)
    $docJson = $wc.DownloadString("http://127.0.0.1:$TestPort/api/document?path=$encodedDocx")
    $docxResp = $docJson | ConvertFrom-Json
    if (-not $docxResp.html -or $docxResp.format -ne "docx") {
        Write-Error "FAIL: GET /api/document failed on DOCX"
    }
    Write-Host "GET /api/document succeeded for both ODT and DOCX."

    # Test 3: POST /api/document (save edited content for both formats)
    $postBodyOdt = @{
        path = $TempOdt
        html = "<h1>Server Test Title ODT</h1><p>Edited via HTTP API</p>"
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $postResultOdt = $wc.UploadString("http://127.0.0.1:$TestPort/api/document", "POST", $postBodyOdt)
    $postRespOdt = $postResultOdt | ConvertFrom-Json
    if (-not $postRespOdt.success) {
        Write-Error "FAIL: POST /api/document failed on ODT"
    }

    $postBodyDocx = @{
        path = $TempDocx
        html = "<h1>Server Test Title DOCX</h1><p>Edited via HTTP API</p>"
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $postResultDocx = $wc.UploadString("http://127.0.0.1:$TestPort/api/document", "POST", $postBodyDocx)
    $postRespDocx = $postResultDocx | ConvertFrom-Json
    if (-not $postRespDocx.success) {
        Write-Error "FAIL: POST /api/document failed on DOCX"
    }
    Write-Host "POST /api/document save succeeded for both ODT and DOCX."

    # Test 4: POST /api/document with createNew: true
    if (Test-Path $CreatedNewDocx) { Remove-Item -Path $CreatedNewDocx -Force }
    $createBody = @{
        createNew = $true
        path = $CreatedNewDocx
        title = "Brand New Document"
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $createResult = $wc.UploadString("http://127.0.0.1:$TestPort/api/document", "POST", $createBody)
    $createResp = $createResult | ConvertFrom-Json
    if (-not $createResp.success -or -not (Test-Path $CreatedNewDocx)) {
        Write-Error "FAIL: POST /api/document with createNew failed"
    }
    Write-Host "POST /api/document createNew succeeded."

    # Test 5: POST /api/convert (cross-conversion and Markdown export)
    if (Test-Path $ConvertedDocx) { Remove-Item -Path $ConvertedDocx -Force }
    $convertBody1 = @{
        sourcePath = $TempOdt
        targetPath = $ConvertedDocx
        format = "docx"
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $convertResult1 = $wc.UploadString("http://127.0.0.1:$TestPort/api/convert", "POST", $convertBody1)
    $convertResp1 = $convertResult1 | ConvertFrom-Json
    if (-not $convertResp1.success -or -not (Test-Path $ConvertedDocx)) {
        Write-Error "FAIL: POST /api/convert ODT -> DOCX failed"
    }

    if (Test-Path $ConvertedMd) { Remove-Item -Path $ConvertedMd -Force }
    $convertBody2 = @{
        html = "<h1>Markdown Heading</h1><p>Paragraph with **bold** idea.</p>"
        targetPath = $ConvertedMd
        format = "md"
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $convertResult2 = $wc.UploadString("http://127.0.0.1:$TestPort/api/convert", "POST", $convertBody2)
    $convertResp2 = $convertResult2 | ConvertFrom-Json
    if (-not $convertResp2.success -or -not (Test-Path $ConvertedMd)) {
        Write-Error "FAIL: POST /api/convert export to Markdown failed"
    }
    Write-Host "POST /api/convert succeeded for both format conversion and Markdown export."

    # Test 6: GET /api/directories
    $dirsJson = $wc.DownloadString("http://127.0.0.1:$TestPort/api/directories")
    $dirsResp = $dirsJson | ConvertFrom-Json
    if (-not $dirsResp.current -or -not $dirsResp.quickLocations) {
        Write-Error "FAIL: GET /api/directories failed"
    }
    Write-Host "GET /api/directories succeeded. Found $($dirsResp.quickLocations.Count) quick locations."

    # Test 7: POST /api/import (base64 with data URL prefix)
    $sampleBytes = [System.IO.File]::ReadAllBytes($TempDocx)
    $b64 = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64," + [System.Convert]::ToBase64String($sampleBytes)
    $importBody = @{
        filename = "imported_test.docx"
        base64 = $b64
    } | ConvertTo-Json
    $wc.Headers.Set("Content-Type", "application/json")
    $importResult = $wc.UploadString("http://127.0.0.1:$TestPort/api/import", "POST", $importBody)
    $importResp = $importResult | ConvertFrom-Json
    if (-not $importResp.path -or -not (Test-Path $importResp.path)) {
        Write-Error "FAIL: POST /api/import failed"
    }
    Remove-Item -Path $importResp.path -Force -ErrorAction SilentlyContinue
    Write-Host "POST /api/import succeeded."

    Write-Host "PASS: Task 2 server tests passed!"
} finally {
    try {
        $wc = New-Object System.Net.WebClient
        $wc.DownloadString("http://127.0.0.1:$TestPort/api/shutdown") | Out-Null
    } catch {}
    if ($job) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -Path $TempOdt, $TempDocx, $SampleOdt, $SampleDocx, $CreatedNewDocx, $ConvertedDocx, $ConvertedMd -Force -ErrorAction SilentlyContinue
}
