# Helper to generate a blank/sample ODT file with generic text
param(
    [string]$OutputPath = "sample.odt"
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (Test-Path $OutputPath) {
    Remove-Item -Path $OutputPath -Force
}

$zip = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Create)

# 1. mimetype (must be first, plain text)
$mimetypeEntry = $zip.CreateEntry("mimetype", [System.IO.Compression.CompressionLevel]::NoCompression)
$mWriter = New-Object System.IO.StreamWriter($mimetypeEntry.Open(), [System.Text.Encoding]::ASCII)
$mWriter.Write("application/vnd.oasis.opendocument.text")
$mWriter.Close()

# 2. META-INF/manifest.xml
$manifestEntry = $zip.CreateEntry("META-INF/manifest.xml")
$manWriter = New-Object System.IO.StreamWriter($manifestEntry.Open(), [System.Text.Encoding]::UTF8)
$manWriter.Write('<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>')
$manWriter.Close()

# 3. content.xml
$contentEntry = $zip.CreateEntry("content.xml")
$cWriter = New-Object System.IO.StreamWriter($contentEntry.Open(), [System.Text.Encoding]::UTF8)
$sampleContent = @'
<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  <office:automatic-styles/>
  <office:body>
    <office:text>
      <text:h text:outline-level="1">Sample Document</text:h>
      <text:p>Welcome to Antigravity ODT Editor. This is a generic sample document.</text:p>
      <text:h text:outline-level="2">1. Overview</text:h>
      <text:p>This document can be used to test viewing, formatting, and live disk saving.</text:p>
    </office:text>
  </office:body>
</office:document-content>
'@
$cWriter.Write($sampleContent.Trim())
$cWriter.Close()

$zip.Dispose()
Write-Host "Created generic sample ODT at $OutputPath"
