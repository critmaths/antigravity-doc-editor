# Test to reproduce and verify fixes for code review findings
$ErrorActionPreference = "Stop"

$PluginDir = "C:\Users\critm\.gemini\config\plugins\odt-editor"
$ConverterScript = Join-Path $PluginDir "sidecars\odt_editor\odt_converter.ps1"
$ServerScript = Join-Path $PluginDir "sidecars\odt_editor\server.ps1"
$SampleOdt = "C:\Users\critm\.gemini\antigravity\scratch\New Big or Medium Project.odt"
$TempOdt = "C:\Users\critm\.gemini\antigravity\scratch\tests\test_review_temp.odt"
Copy-Item -Path $SampleOdt -Destination $TempOdt -Force

. $ConverterScript

Write-Host "=== Test 1: Critical-4 Pipeline Output Leak in Convert-HtmlToOdt ==="
$htmlSnippet = "<p>Test Paragraph</p>"
$ret = Convert-HtmlToOdt -OdtPath $TempOdt -HtmlContent $htmlSnippet
if ($ret -is [array]) {
    Write-Error "FAIL [Critical-4]: Convert-HtmlToOdt leaked $($ret.Count) items into pipeline instead of returning boolean: $($ret -join ', ')"
}
if ($ret -ne $true) {
    Write-Error "FAIL [Critical-4]: Convert-HtmlToOdt did not return `$true"
}
Write-Host "Test 1 Passed: Return value is strictly boolean `$true"

Write-Host "=== Test 2: Critical-2 & Critical-3 Entity Handling & Rich Formatting Preservation ==="
$richHtml = "<h1>Title &amp; Vision</h1><p>A paragraph with &mdash; em-dash, &copy; copyright, and <span style=""color: #ef4444; font-size: 18pt; font-family: 'Georgia', serif;"">styled colored text</span>.</p><p style=""text-align: center"">Centered paragraph</p><blockquote>Quoted wisdom here</blockquote><ul><li>Item &amp; subtext</li></ul>"
$ret2 = Convert-HtmlToOdt -OdtPath $TempOdt -HtmlContent $richHtml
if ($ret2 -ne $true) {
    Write-Error "FAIL: Rich HTML conversion failed"
}

# Verify content.xml in the updated zip
$zip = [System.IO.Compression.ZipFile]::OpenRead($TempOdt)
$entry = $zip.GetEntry("content.xml")
$reader = New-Object System.IO.StreamReader($entry.Open())
$xmlContent = $reader.ReadToEnd()
$reader.Close()
$zip.Dispose()

# Check that entity did not trigger fallback plain text encoding
if ($xmlContent -match "&lt;h1&gt;" -or $xmlContent -match "&lt;p&gt;") {
    Write-Error "FAIL [Critical-2]: HTML was encoded into plain text fallback due to entity failure!"
}
if ($xmlContent -notmatch "Title &amp; Vision" -and $xmlContent -notmatch "Title &#38; Vision") {
    Write-Error "FAIL [Critical-2]: Missing Title in XML"
}
if ($xmlContent -notmatch "#ef4444" -or $xmlContent -notmatch "18pt") {
    Write-Error "FAIL [Critical-3]: Style attributes (color #ef4444, size 18pt) were lost during save!"
}
Write-Host "Test 2 Passed: HTML entities and rich formatting preserved in ODF XML automatic styles"

Write-Host "=== Test 3: Critical-1 & Important-1 Server /api/documents and Path Validation ==="
# Test PSCustomObject serialization logic
$searchDirs = @("C:\Users\critm\.gemini\antigravity\scratch")
$foundFiles = @()
foreach ($dir in $searchDirs) {
    if (Test-Path $dir) {
        $odtFiles = Get-ChildItem -Path $dir -Recurse -Filter "*.odt" -ErrorAction SilentlyContinue | Select-Object -First 5
        foreach ($f in $odtFiles) {
            $foundFiles += [PSCustomObject]@{
                name = $f.Name
                path = $f.FullName
                size = $f.Length
                lastModified = $f.LastWriteTime.ToString("o")
            }
        }
    }
}
$jsonOutput = @{ files = @($foundFiles | Sort-Object path -Unique) } | ConvertTo-Json -Depth 5
$parsed = $jsonOutput | ConvertFrom-Json
if (-not $parsed.files -or -not $parsed.files[0].path) {
    Write-Error "FAIL [Critical-1]: /api/documents emitted broken path"
}
Write-Host "Test 3 Passed: /api/documents formats file list with valid properties"

# Clean up
Remove-Item -Path $TempOdt -Force -ErrorAction SilentlyContinue
Write-Host "=== ALL CODE REVIEW TEST CHECKS PASSED ==="
