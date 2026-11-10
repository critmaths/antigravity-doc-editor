# Test for server.ps1
$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$ServerScript = Join-Path $RepoDir "sidecars\odt_editor\server.ps1"
$CreateSampleScript = Join-Path $ScriptDir "create_sample.ps1"
$SampleOdt = Join-Path $ScriptDir "sample.odt"
$TempOdt = Join-Path $ScriptDir "server_test_temp.odt"

& $CreateSampleScript -OutputPath $SampleOdt
Copy-Item -Path $SampleOdt -Destination $TempOdt -Force

Write-Host "Checking if server script exists: $ServerScript"
if (-not (Test-Path $ServerScript)) {
    Write-Error "FAIL: $ServerScript does not exist"
}

$TestPort = 58925
$job = Start-Job -ScriptBlock {
    param($s, $p)
    & $s -Port $p
} -ArgumentList $ServerScript, $TestPort

try {
    # Wait for server to start listening
    $started = $false
    for ($i = 0; $i -lt 10; $i++) {
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
    Write-Host "GET /api/documents returned $($docsResp.files.Count) files"

    # Test 2: GET /api/document?path=...
    $encodedPath = [System.Net.WebUtility]::UrlEncode($TempOdt)
    $docJson = $wc.DownloadString("http://127.0.0.1:$TestPort/api/document?path=$encodedPath")
    $docResp = $docJson | ConvertFrom-Json
    if (-not $docResp.html -or $docResp.html.Length -eq 0) {
        Write-Error "FAIL: /api/document returned empty html"
    }
    Write-Host "GET /api/document returned HTML of length $($docResp.html.Length)"

    # Test 3: POST /api/document
    $postBody = @{
        path = $TempOdt
        html = "<h1>Server Test Title</h1><p>Edited via HTTP API</p>"
    } | ConvertTo-Json
    $wc.Headers.Add("Content-Type", "application/json")
    $postResult = $wc.UploadString("http://127.0.0.1:$TestPort/api/document", "POST", $postBody)
    $postResp = $postResult | ConvertFrom-Json
    if (-not $postResp.success) {
        Write-Error "FAIL: POST /api/document did not return success"
    }
    Write-Host "POST /api/document returned success"

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
    Remove-Item -Path $TempOdt, $SampleOdt -Force -ErrorAction SilentlyContinue
}
