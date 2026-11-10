# Test for odt_converter.ps1
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$ConverterScript = Join-Path $RepoDir "sidecars\odt_editor\odt_converter.ps1"
$CreateSampleScript = Join-Path $ScriptDir "create_sample.ps1"
$SampleOdt = Join-Path $ScriptDir "sample.odt"
$TempOdt = Join-Path $ScriptDir "temp_test.odt"

# Generate fresh generic sample
& $CreateSampleScript -OutputPath $SampleOdt

Write-Host "Checking if converter exists: $ConverterScript"
if (-not (Test-Path $ConverterScript)) {
    Write-Error "FAIL: $ConverterScript does not exist"
}

. $ConverterScript

Write-Host "Testing Convert-OdtToHtml..."
$doc = Convert-OdtToHtml -OdtPath $SampleOdt
if (-not $doc) {
    Write-Error "FAIL: Convert-OdtToHtml returned null"
}
if (-not $doc.html -or $doc.html.Length -eq 0) {
    Write-Error "FAIL: Document HTML is empty"
}
Write-Host "HTML length extracted: $($doc.html.Length)"

# Test round-trip serialization
Copy-Item -Path $SampleOdt -Destination $TempOdt -Force

$testHtml = "<h1>Project Vision</h1><p>This is a <b>bold</b> test paragraph edited by Antigravity ODT Editor.</p><ul><li>First Item</li><li>Second Item</li></ul>"
Write-Host "Testing Convert-HtmlToOdt..."
$success = Convert-HtmlToOdt -OdtPath $TempOdt -HtmlContent $testHtml
if ($success -ne $true) {
    Write-Error "FAIL: Convert-HtmlToOdt failed"
}

# Verify content.xml inside the updated zip
$zip = [System.IO.Compression.ZipFile]::OpenRead($TempOdt)
$entry = $zip.GetEntry("content.xml")
if (-not $entry) {
    $zip.Dispose()
    Write-Error "FAIL: content.xml missing from updated zip"
}
$reader = New-Object System.IO.StreamReader($entry.Open())
$xmlContent = $reader.ReadToEnd()
$reader.Close()
$zip.Dispose()

if ($xmlContent -notmatch "Project Vision" -or $xmlContent -notmatch "Antigravity ODT Editor") {
    Write-Error "FAIL: Updated content.xml does not contain expected text"
}

# Clean up temp files
Remove-Item -Path $TempOdt, $SampleOdt -Force -ErrorAction SilentlyContinue

Write-Host "PASS: Task 1 converter tests succeeded!"
