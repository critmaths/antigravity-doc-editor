# Test for doc_converter.ps1 (ODT + DOCX + Markdown + Blank creation + Cross-conversion)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$ConverterScript = Join-Path $RepoDir "sidecars\doc_editor\doc_converter.ps1"
$CreateSampleScript = Join-Path $ScriptDir "create_sample.ps1"
$SampleOdt = Join-Path $ScriptDir "sample.odt"
$SampleDocx = Join-Path $ScriptDir "sample.docx"
$TempOdt = Join-Path $ScriptDir "temp_test.odt"
$TempDocx = Join-Path $ScriptDir "temp_test.docx"
$ConvertedDocx = Join-Path $ScriptDir "converted_from_odt.docx"
$ConvertedOdt = Join-Path $ScriptDir "converted_from_docx.odt"
$ConvertedMd = Join-Path $ScriptDir "converted_export.md"
$BlankDocx = Join-Path $ScriptDir "blank_created.docx"

Write-Host "Checking if converter exists: $ConverterScript"
if (-not (Test-Path $ConverterScript)) {
    Write-Error "FAIL: $ConverterScript does not exist"
}

. $ConverterScript
& $CreateSampleScript -OutputPath $SampleOdt
& $CreateSampleScript -OutputPath $SampleDocx

try {
    Write-Host "=== 1. Testing Convert-DocumentToHtml on ODT and DOCX ==="
    $odtDoc = Convert-DocumentToHtml -DocPath $SampleOdt
    if (-not $odtDoc -or $odtDoc.format -ne "odt" -or $odtDoc.html -notmatch "Sample Document") {
        Write-Error "FAIL: Convert-DocumentToHtml failed on ODT"
    }
    $docxDoc = Convert-DocumentToHtml -DocPath $SampleDocx
    if (-not $docxDoc -or $docxDoc.format -ne "docx" -or $docxDoc.html -notmatch "Sample Document") {
        Write-Error "FAIL: Convert-DocumentToHtml failed on DOCX"
    }
    Write-Host "Both ODT ($($odtDoc.html.Length) chars) and DOCX ($($docxDoc.html.Length) chars) parsed to HTML."

    Write-Host "=== 2. Testing Rich Formatting Round-Trip on ODT and DOCX ==="
    Copy-Item -Path $SampleOdt -Destination $TempOdt -Force
    Copy-Item -Path $SampleDocx -Destination $TempDocx -Force

    $richHtml = @"
<h1>Project Vision &amp; Strategy</h1>
<h2>Section 2: Formatting &mdash; Details</h2>
<h3>Subsection 3</h3>
<p style="text-align: center">Centered paragraph with <b>bold</b>, <i>italic</i>, <u>underline</u>, and <s>strikethrough</s> text.</p>
<p><span style="color: #dc2626; background-color: #fef08a; font-size: 18pt; font-family: 'Georgia', serif;">Colored highlighted styled text</span> and <font color="#2563eb" face="Calibri" size="5">legacy font tag</font>.</p>
<blockquote>Quoted callout block with important notes.</blockquote>
<ul><li>Bullet One</li><li>Bullet Two</li></ul>
<ol><li>Numbered First</li><li>Numbered Second</li></ol>
<p>Visit <a href="https://github.com/critmaths">Project Repository</a> for details.</p>
<table><tr><th>Header A</th><th>Header B</th></tr><tr><td>Cell 1</td><td>Cell 2</td></tr></table>
"@

    $retOdt = Convert-HtmlToDocument -DocPath $TempOdt -HtmlContent $richHtml
    if ($retOdt -is [array] -or $retOdt -ne $true) {
        Write-Error "FAIL: Convert-HtmlToDocument on ODT failed or leaked pipeline output: $retOdt"
    }

    $retDocx = Convert-HtmlToDocument -DocPath $TempDocx -HtmlContent $richHtml
    if ($retDocx -is [array] -or $retDocx -ne $true) {
        Write-Error "FAIL: Convert-HtmlToDocument on DOCX failed or leaked pipeline output: $retDocx"
    }

    # Verify DOCX package integrity (sectPr, numbering.xml, document.xml.rels, styles)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($TempDocx)
    try {
        $docEntry = $zip.GetEntry("word/document.xml")
        $numEntry = $zip.GetEntry("word/numbering.xml")
        $relsEntry = $zip.GetEntry("word/_rels/document.xml.rels")
        if (-not $docEntry -or -not $numEntry -or -not $relsEntry) {
            Write-Error "FAIL: DOCX missing required OpenXML entries (document.xml, numbering.xml, or document.xml.rels)"
        }
        $r = New-Object System.IO.StreamReader($docEntry.Open())
        $docXml = $r.ReadToEnd()
        $r.Close()
        if ($docXml -notmatch "w:sectPr") {
            Write-Error "FAIL: DOCX lost w:sectPr section properties!"
        }
        if ($docXml -notmatch "dc2626" -or $docXml -notmatch "w:hyperlink" -or $docXml -notmatch "w:tbl") {
            Write-Error "FAIL: DOCX missing color, hyperlink, or table elements in word/document.xml"
        }
    } finally {
        $zip.Dispose()
    }

    # Verify round-trip read back from both ODT and DOCX
    $readBackOdt = Convert-DocumentToHtml -DocPath $TempOdt
    $readBackDocx = Convert-DocumentToHtml -DocPath $TempDocx
    foreach ($rb in @($readBackOdt, $readBackDocx)) {
        if ($rb.html -notmatch "Project Vision" -or $rb.html -notmatch "Bullet One" -or $rb.html -notmatch "Cell 1" -or $rb.html -notmatch "https://github.com/critmaths") {
            Write-Error "FAIL: Round-trip HTML readback for $($rb.format) lost content: $($rb.html)"
        }
    }
    Write-Host "Rich formatting round-trip verified for both ODT and DOCX."

    Write-Host "=== 3. Testing Cross-Format Conversion (ODT <-> DOCX & Markdown) ==="
    $conv1 = Convert-DocumentFormat -SourcePath $TempOdt -TargetPath $ConvertedDocx -HtmlContent $richHtml
    if ($conv1 -ne $true -or -not (Test-Path $ConvertedDocx)) {
        Write-Error "FAIL: ODT -> DOCX conversion failed"
    }
    $checkDocx = Convert-DocumentToHtml -DocPath $ConvertedDocx
    if ($checkDocx.format -ne "docx" -or $checkDocx.html -notmatch "Project Vision") {
        Write-Error "FAIL: Converted DOCX failed readback verification"
    }

    $conv2 = Convert-DocumentFormat -SourcePath $TempDocx -TargetPath $ConvertedOdt -HtmlContent $richHtml
    if ($conv2 -ne $true -or -not (Test-Path $ConvertedOdt)) {
        Write-Error "FAIL: DOCX -> ODT conversion failed"
    }
    $checkOdt = Convert-DocumentToHtml -DocPath $ConvertedOdt
    if ($checkOdt.format -ne "odt" -or $checkOdt.html -notmatch "Project Vision") {
        Write-Error "FAIL: Converted ODT failed readback verification"
    }

    $conv3 = Convert-DocumentFormat -SourcePath $TempDocx -TargetPath $ConvertedMd -HtmlContent $richHtml
    if ($conv3 -ne $true -or -not (Test-Path $ConvertedMd)) {
        Write-Error "FAIL: Export to Markdown failed"
    }
    $mdText = Get-Content $ConvertedMd -Raw
    if ($mdText -notmatch "# Project Vision" -or $mdText -notmatch "Project Repository") {
        Write-Error "FAIL: Exported Markdown missing expected content"
    }

    $ConvertedHtml = Join-Path $ScriptDir "converter_test.html"
    $ConvertedTxt = Join-Path $ScriptDir "converter_test.txt"
    $conv4 = Convert-DocumentFormat -SourcePath $TempDocx -TargetPath $ConvertedHtml -HtmlContent $richHtml
    if ($conv4 -ne $true -or -not (Test-Path $ConvertedHtml)) {
        Write-Error "FAIL: Export to HTML failed"
    }
    $htmlText = Get-Content $ConvertedHtml -Raw
    if ($htmlText -notmatch "<!DOCTYPE html>" -or $htmlText -notmatch "Project Vision") {
        Write-Error "FAIL: Exported HTML missing expected content"
    }

    $conv5 = Convert-DocumentFormat -SourcePath $TempDocx -TargetPath $ConvertedTxt -HtmlContent $richHtml
    if ($conv5 -ne $true -or -not (Test-Path $ConvertedTxt)) {
        Write-Error "FAIL: Export to Plain Text failed"
    }
    $plainText = Get-Content $ConvertedTxt -Raw
    if ($plainText -notmatch "Project Vision") {
        Write-Error "FAIL: Exported Plain Text missing expected content"
    }
    Remove-Item -Path $ConvertedHtml, $ConvertedTxt -Force -ErrorAction SilentlyContinue

    Write-Host "=== 4. Testing New-BlankDocument ==="
    $blankCreated = New-BlankDocument -DocPath $BlankDocx -Title "Fresh Spec"
    if (-not $blankCreated -or -not (Test-Path $BlankDocx)) {
        Write-Error "FAIL: New-BlankDocument failed for DOCX"
    }
    $blankRead = Convert-DocumentToHtml -DocPath $BlankDocx
    if ($blankRead.html -notmatch "Fresh Spec") {
        Write-Error "FAIL: New-BlankDocument did not initialize title heading"
    }

    Write-Host "PASS: Task 1 unified converter tests succeeded!"
} finally {
    Remove-Item -Path $SampleOdt, $SampleDocx, $TempOdt, $TempDocx, $ConvertedDocx, $ConvertedOdt, $ConvertedMd, $BlankDocx -Force -ErrorAction SilentlyContinue
}
