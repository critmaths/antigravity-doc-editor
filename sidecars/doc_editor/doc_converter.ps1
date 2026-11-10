<#
.SYNOPSIS
    Unified Office Document Converter for Antigravity Document Editor (doc-editor).
    Provides bidirectional conversion between OpenDocument Text (.odt),
    Microsoft Word OpenXML (.docx), HTML, and Markdown.
#>

param(
    [string]$Read = "",
    [ValidateSet("html", "markdown", "text")]
    [string]$Format = "markdown"
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ZipEntryText {
    param(
        [string]$ZipPath,
        [string]$EntryName
    )
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $entry = $zip.GetEntry($EntryName)
        if ($null -eq $entry) { return $null }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        try {
            return $reader.ReadToEnd()
        } finally {
            $reader.Close()
        }
    } finally {
        $zip.Dispose()
    }
}

function Set-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName,
        [string]$Content,
        [System.IO.Compression.CompressionLevel]$Level = [System.IO.Compression.CompressionLevel]::Optimal
    )
    if ($Zip.Mode -ne [System.IO.Compression.ZipArchiveMode]::Create) {
        $existing = $Zip.GetEntry($EntryName)
        if ($existing) { $existing.Delete() }
    }
    $entry = $Zip.CreateEntry($EntryName, $Level)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), $utf8NoBom)
    try {
        $writer.Write($Content)
    } finally {
        $writer.Close()
    }
}

function New-BlankOdt {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OdtPath,
        [string]$Title = "Untitled Document"
    )
    if (Test-Path $OdtPath) { Remove-Item -Path $OdtPath -Force }
    $parentDir = Split-Path -Parent $OdtPath
    if ($parentDir -and -not (Test-Path $parentDir)) {
        [void](New-Item -ItemType Directory -Path $parentDir -Force)
    }

    $zip = [System.IO.Compression.ZipFile]::Open($OdtPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $mimetypeEntry = $zip.CreateEntry("mimetype", [System.IO.Compression.CompressionLevel]::NoCompression)
        $mWriter = New-Object System.IO.StreamWriter($mimetypeEntry.Open(), [System.Text.Encoding]::ASCII)
        $mWriter.Write("application/vnd.oasis.opendocument.text")
        $mWriter.Close()

        $manifestXml = '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>'
        Set-ZipEntryText -Zip $zip -EntryName "META-INF/manifest.xml" -Content $manifestXml

        $safeTitle = [System.Net.WebUtility]::HtmlEncode($Title)
        $contentXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
  <office:automatic-styles/>
  <office:body>
    <office:text>
      <text:h text:outline-level="1">$safeTitle</text:h>
      <text:p></text:p>
    </office:text>
  </office:body>
</office:document-content>
"@
        Set-ZipEntryText -Zip $zip -EntryName "content.xml" -Content $contentXml.Trim()
    } finally {
        $zip.Dispose()
    }
    return $true
}

function Get-DefaultDocxStylesXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:sz w:val="24"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="48"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Quote">
    <w:name w:val="Quote"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="720"/></w:pPr>
    <w:rPr><w:i/><w:color w:val="4B5563"/></w:rPr>
  </w:style>
</w:styles>
'@.Trim()
}

function Get-DefaultDocxNumberingXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>
'@.Trim()
}

function New-BlankDocx {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DocxPath,
        [string]$Title = "Untitled Document"
    )
    if (Test-Path $DocxPath) { Remove-Item -Path $DocxPath -Force }
    $parentDir = Split-Path -Parent $DocxPath
    if ($parentDir -and -not (Test-Path $parentDir)) {
        [void](New-Item -ItemType Directory -Path $parentDir -Force)
    }

    $zip = [System.IO.Compression.ZipFile]::Open($DocxPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $contentTypesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
'@.Trim()
        Set-ZipEntryText -Zip $zip -EntryName "[Content_Types].xml" -Content $contentTypesXml

        $rootRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@.Trim()
        Set-ZipEntryText -Zip $zip -EntryName "_rels/.rels" -Content $rootRelsXml

        $docRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
'@.Trim()
        Set-ZipEntryText -Zip $zip -EntryName "word/_rels/document.xml.rels" -Content $docRelsXml

        Set-ZipEntryText -Zip $zip -EntryName "word/styles.xml" -Content (Get-DefaultDocxStylesXml)
        Set-ZipEntryText -Zip $zip -EntryName "word/numbering.xml" -Content (Get-DefaultDocxNumberingXml)

        $safeTitle = [System.Net.WebUtility]::HtmlEncode($Title)
        $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>$safeTitle</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t></w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
"@.Trim()
        Set-ZipEntryText -Zip $zip -EntryName "word/document.xml" -Content $documentXml
    } finally {
        $zip.Dispose()
    }
    return $true
}

function New-BlankDocument {
    param(
        [Parameter(Mandatory=$true)]
        [Alias("Path")]
        [string]$DocPath,
        [string]$Title = "Untitled Document"
    )
    $ext = [System.IO.Path]::GetExtension($DocPath).ToLower()
    if ($ext -eq ".docx") {
        return (New-BlankDocx -DocxPath $DocPath -Title $Title)
    } else {
        return (New-BlankOdt -OdtPath $DocPath -Title $Title)
    }
}

# ==============================================================================
# ODT PARSER & SERIALIZER
# ==============================================================================

function Parse-OdtAutomaticStyles {
    param([xml]$ContentXml)

    $stylesMap = @{}
    $ns = New-Object System.Xml.XmlNamespaceManager($ContentXml.NameTable)
    $ns.AddNamespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
    $ns.AddNamespace("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
    $ns.AddNamespace("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")

    $styleNodes = $ContentXml.SelectNodes("//office:automatic-styles/style:style", $ns)
    foreach ($node in $styleNodes) {
        $name = $node.GetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        if (-not $name) { continue }

        $cssProps = @()
        $textProps = $node.SelectSingleNode("style:text-properties", $ns)
        if ($textProps) {
            if ($textProps.GetAttribute("font-weight", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0") -eq "bold") {
                $cssProps += "font-weight: bold"
            }
            if ($textProps.GetAttribute("font-style", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0") -eq "italic") {
                $cssProps += "font-style: italic"
            }
            if ($textProps.GetAttribute("text-underline-style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0") -or
                $textProps.GetAttribute("text-decoration", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0") -eq "underline") {
                $cssProps += "text-decoration: underline"
            }
            if ($textProps.GetAttribute("text-line-through-style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")) {
                $cssProps += "text-decoration: line-through"
            }
            $color = $textProps.GetAttribute("color", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
            if ($color) { $cssProps += "color: $color" }
            $bg = $textProps.GetAttribute("background-color", "urn:oasis:names:tc:opendocument:xmlns:fo-compatible:1.0")
            if (-not $bg) { $bg = $textProps.GetAttribute("background-color", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0") }
            if ($bg) { $cssProps += "background-color: $bg" }
            $fontSize = $textProps.GetAttribute("font-size", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
            if ($fontSize) { $cssProps += "font-size: $fontSize" }
            $fontName = $textProps.GetAttribute("font-name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
            if (-not $fontName) { $fontName = $textProps.GetAttribute("font-name", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0") }
            if ($fontName) { $cssProps += "font-family: '$fontName', sans-serif" }
        }

        $parProps = $node.SelectSingleNode("style:paragraph-properties", $ns)
        if ($parProps) {
            $align = $parProps.GetAttribute("text-align", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
            if ($align) { $cssProps += "text-align: $align" }
        }

        if ($cssProps.Count -gt 0) {
            $stylesMap[$name] = ($cssProps -join "; ")
        }
    }

    return $stylesMap
}

function Convert-OdfNodeToHtml {
    param(
        [System.Xml.XmlNode]$Node,
        [hashtable]$StylesMap,
        [System.Xml.XmlNamespaceManager]$Ns
    )

    if ($Node.NodeType -eq [System.Xml.XmlNodeType]::Text) {
        return [System.Net.WebUtility]::HtmlEncode($Node.Value)
    }

    $localName = $Node.LocalName
    $prefix = $Node.Prefix

    if ($prefix -eq "text") {
        switch ($localName) {
            "h" {
                $level = $Node.GetAttribute("outline-level", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                if (-not $level -or [int]$level -gt 4) { $level = 1 }
                $styleName = $Node.GetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                $styleAttr = if ($StylesMap.ContainsKey($styleName)) { " style=""$($StylesMap[$styleName])""" } else { "" }
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                }
                return "<h$level$styleAttr>$innerHtml</h$level>"
            }
            "p" {
                $styleName = $Node.GetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                if ($styleName -eq "P_Quote") {
                    $innerHtml = ""
                    foreach ($child in $Node.ChildNodes) {
                        $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                    }
                    return "<blockquote>$innerHtml</blockquote>"
                }
                $styleAttr = if ($StylesMap.ContainsKey($styleName)) { " style=""$($StylesMap[$styleName])""" } else { "" }
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                }
                if ([string]::IsNullOrWhiteSpace($innerHtml)) { $innerHtml = "<br/>" }
                return "<p$styleAttr>$innerHtml</p>"
            }
            "span" {
                $styleName = $Node.GetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                }
                if ($styleName -eq "T_Bold") { return "<b>$innerHtml</b>" }
                if ($styleName -eq "T_Italic") { return "<i>$innerHtml</i>" }
                if ($styleName -eq "T_Underline") { return "<u>$innerHtml</u>" }
                if ($styleName -eq "T_Strike") { return "<s>$innerHtml</s>" }
                $styleAttr = if ($StylesMap.ContainsKey($styleName)) { " style=""$($StylesMap[$styleName])""" } else { "" }
                return "<span$styleAttr>$innerHtml</span>"
            }
            "list" {
                $styleName = $Node.GetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                $tag = if ($styleName -eq "L_Ordered") { "ol" } else { "ul" }
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                }
                return "<$tag>$innerHtml</$tag>"
            }
            "list-item" {
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    if ($child.LocalName -eq "p") {
                        foreach ($grandChild in $child.ChildNodes) {
                            $innerHtml += Convert-OdfNodeToHtml -Node $grandChild -StylesMap $StylesMap -Ns $Ns
                        }
                    } else {
                        $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                    }
                }
                return "<li>$innerHtml</li>"
            }
            "line-break" { return "<br/>" }
            "tab" { return "&emsp;" }
            "s" {
                $count = 1
                $c = $Node.GetAttribute("c", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                if ($c) { [int]::TryParse($c, [ref]$count) | Out-Null }
                return ("&nbsp;" * $count)
            }
            "a" {
                $href = $Node.GetAttribute("href", "http://www.w3.org/1999/xlink")
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                }
                return "<a href=""$href"">$innerHtml</a>"
            }
        }
    }

    if ($prefix -eq "table" -and $localName -eq "table") {
        $innerHtml = ""
        foreach ($child in $Node.ChildNodes) {
            $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
        }
        return "<table class=""doc-table"">$innerHtml</table>"
    }
    if ($prefix -eq "table" -and $localName -eq "table-row") {
        $innerHtml = ""
        foreach ($child in $Node.ChildNodes) {
            $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
        }
        return "<tr>$innerHtml</tr>"
    }
    if ($prefix -eq "table" -and $localName -eq "table-cell") {
        $innerHtml = ""
        foreach ($child in $Node.ChildNodes) {
            if ($child.LocalName -eq "p") {
                foreach ($gc in $child.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $gc -StylesMap $StylesMap -Ns $Ns
                }
            } else {
                $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
            }
        }
        return "<td>$innerHtml</td>"
    }

    $fallbackHtml = ""
    foreach ($child in $Node.ChildNodes) {
        $fallbackHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
    }
    return $fallbackHtml
}

function Convert-OdtToHtml {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OdtPath
    )
    if (-not (Test-Path $OdtPath)) { throw "File not found: $OdtPath" }

    $contentXmlStr = Get-ZipEntryText -ZipPath $OdtPath -EntryName "content.xml"
    if (-not $contentXmlStr) { throw "Missing content.xml in ODT file: $OdtPath" }

    [xml]$xml = $contentXmlStr
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
    $ns.AddNamespace("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
    $ns.AddNamespace("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
    $ns.AddNamespace("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
    $ns.AddNamespace("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")

    $stylesMap = Parse-OdtAutomaticStyles -ContentXml $xml
    $bodyNode = $xml.SelectSingleNode("//office:body/office:text", $ns)
    if (-not $bodyNode) {
        return @{
            title = [System.IO.Path]::GetFileNameWithoutExtension($OdtPath)
            format = "odt"
            html = "<p></p>"
            styles = ""
        }
    }

    $htmlParts = @()
    foreach ($child in $bodyNode.ChildNodes) {
        $htmlParts += Convert-OdfNodeToHtml -Node $child -StylesMap $stylesMap -Ns $ns
    }

    return @{
        title = [System.IO.Path]::GetFileNameWithoutExtension($OdtPath)
        format = "odt"
        html = ($htmlParts -join "`n")
        styles = ""
    }
}

function Normalize-HtmlForXml {
    param([string]$html)
    $clean = $html.Replace("&nbsp;", " ").Replace("&emsp;", "    ")
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "<br\s*>", "<br/>")
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "<hr\s*>", "<hr/>")
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "<img([^>]*?)(?<!/)>", "<img`$1/>")

    $entityCallback = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $ent = $m.Groups[1].Value.ToLower()
        if ($ent -in @("amp", "lt", "gt", "quot", "apos")) { return $m.Value }
        $decoded = [System.Net.WebUtility]::HtmlDecode($m.Value)
        if ($decoded -eq "&") { return "&amp;" }
        if ($decoded -eq "<") { return "&lt;" }
        if ($decoded -eq ">") { return "&gt;" }
        if ($decoded -eq '"') { return "&quot;" }
        if ($decoded -eq "'") { return "&apos;" }
        return $decoded
    }
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "&([a-zA-Z0-9]+);", $entityCallback)
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "&(?!(amp|lt|gt|quot|apos|#\d+|#[xX][0-9a-fA-F]+);)", "&amp;")
    return $clean
}

function Convert-HtmlToOdt {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OdtPath,
        [Parameter(Mandatory=$true)]
        [string]$HtmlContent
    )

    if (-not (Test-Path $OdtPath)) {
        [void](New-BlankOdt -OdtPath $OdtPath)
    }

    $existingXmlStr = Get-ZipEntryText -ZipPath $OdtPath -EntryName "content.xml"
    [xml]$xml = $existingXmlStr
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
    $ns.AddNamespace("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
    $ns.AddNamespace("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
    $ns.AddNamespace("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
    $ns.AddNamespace("xlink", "http://www.w3.org/1999/xlink")
    $ns.AddNamespace("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")

    $autoStylesNode = $xml.SelectSingleNode("//office:automatic-styles", $ns)
    if (-not $autoStylesNode) {
        $autoStylesNode = $xml.CreateElement("office", "automatic-styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        [void]$xml.DocumentElement.PrependChild($autoStylesNode)
    }

    $helperStyles = @(
        @{ name = "T_Bold"; family = "text"; prop = "fo:font-weight"; val = "bold" },
        @{ name = "T_Italic"; family = "text"; prop = "fo:font-style"; val = "italic" },
        @{ name = "T_Underline"; family = "text"; prop = "style:text-underline-style"; val = "solid" },
        @{ name = "T_Strike"; family = "text"; prop = "style:text-line-through-style"; val = "solid" },
        @{ name = "P_Quote"; family = "paragraph"; prop = "fo:margin-left"; val = "0.5in" }
    )
    foreach ($hs in $helperStyles) {
        $existing = $autoStylesNode.SelectSingleNode("style:style[@style:name='$($hs.name)']", $ns)
        if (-not $existing) {
            $sEl = $xml.CreateElement("style", "style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
            [void]$sEl.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", $hs.name)
            [void]$sEl.SetAttribute("family", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", $hs.family)
            $propTag = if ($hs.family -eq "text") { "style:text-properties" } else { "style:paragraph-properties" }
            $tp = $xml.CreateElement("style", $propTag.Split(":")[1], "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
            $parts = $hs.prop.Split(":")
            $nsUri = if ($parts[0] -eq "fo") { "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" } else { "urn:oasis:names:tc:opendocument:xmlns:style:1.0" }
            [void]$tp.SetAttribute($parts[1], $nsUri, $hs.val)
            [void]$sEl.AppendChild($tp)
            [void]$autoStylesNode.AppendChild($sEl)
        }
    }

    $bodyTextNode = $xml.SelectSingleNode("//office:body/office:text", $ns)
    if (-not $bodyTextNode) {
        $bodyNode = $xml.SelectSingleNode("//office:body", $ns)
        if (-not $bodyNode) {
            $bodyNode = $xml.CreateElement("office", "body", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
            [void]$xml.DocumentElement.AppendChild($bodyNode)
        }
        $bodyTextNode = $xml.CreateElement("office", "text", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        [void]$bodyNode.AppendChild($bodyTextNode)
    }

    [void]$bodyTextNode.RemoveAll()

    $normalizedHtml = Normalize-HtmlForXml -html $HtmlContent
    $wrappedHtml = "<root>" + $normalizedHtml + "</root>"

    $htmlDoc = New-Object System.Xml.XmlDocument
    try {
        $htmlDoc.LoadXml($wrappedHtml)
    } catch {
        $safeText = [System.Net.WebUtility]::HtmlEncode($HtmlContent)
        $htmlDoc.LoadXml("<root><p>$safeText</p></root>")
    }

    $script:dynStyleId = 100
    function Get-OrCreateDynamicStyle {
        param([string]$Family, [hashtable]$TextProps, [hashtable]$ParProps)
        $script:dynStyleId++
        $styleName = "DS_$($script:dynStyleId)"
        $sEl = $xml.CreateElement("style", "style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        [void]$sEl.SetAttribute("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", $styleName)
        [void]$sEl.SetAttribute("family", "urn:oasis:names:tc:opendocument:xmlns:style:1.0", $Family)

        if ($TextProps.Count -gt 0) {
            $tp = $xml.CreateElement("style", "text-properties", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
            foreach ($k in $TextProps.Keys) {
                [void]$tp.SetAttribute($k, "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0", $TextProps[$k])
            }
            [void]$sEl.AppendChild($tp)
        }
        if ($ParProps.Count -gt 0) {
            $pp = $xml.CreateElement("style", "paragraph-properties", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
            foreach ($k in $ParProps.Keys) {
                [void]$pp.SetAttribute($k, "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0", $ParProps[$k])
            }
            [void]$sEl.AppendChild($pp)
        }

        [void]$autoStylesNode.AppendChild($sEl)
        return $styleName
    }

    function Import-HtmlToOdf {
        param([System.Xml.XmlNode]$HNode, [System.Xml.XmlNode]$OdfTarget, [System.Xml.XmlDocument]$TargetDoc)

        foreach ($child in $HNode.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Text) {
                $textNode = $TargetDoc.CreateTextNode($child.Value)
                [void]$OdfTarget.AppendChild($textNode)
                continue
            }

            $tag = $child.LocalName.ToLower()
            $styleAttr = if ($child.Attributes -and $child.Attributes["style"]) { $child.Attributes["style"].Value } else { "" }
            $textProps = @{}
            $parProps = @{}
            if ($styleAttr) {
                foreach ($pair in $styleAttr.Split(";")) {
                    $kv = $pair.Split(":")
                    if ($kv.Length -eq 2) {
                        $propName = $kv[0].Trim().ToLower()
                        $propVal = $kv[1].Trim().Trim("'", '"')
                        if ($propName -eq "color") { $textProps["color"] = $propVal }
                        elseif ($propName -eq "background-color") { $textProps["background-color"] = $propVal }
                        elseif ($propName -eq "font-size") { $textProps["font-size"] = $propVal }
                        elseif ($propName -eq "font-family") { $textProps["font-name"] = $propVal.Split(",")[0].Trim() }
                        elseif ($propName -eq "font-weight" -and $propVal -eq "bold") { $textProps["font-weight"] = "bold" }
                        elseif ($propName -eq "font-style" -and $propVal -eq "italic") { $textProps["font-style"] = "italic" }
                        elseif ($propName -eq "text-align") { $parProps["text-align"] = $propVal }
                    }
                }
            }

            switch ($tag) {
                { $_ -match "^h[1-6]$" } {
                    $level = $_.Substring(1)
                    $hEl = $TargetDoc.CreateElement("text", "h", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$hEl.SetAttribute("outline-level", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", $level)
                    if ($parProps.Count -gt 0 -or $textProps.Count -gt 0) {
                        $sName = Get-OrCreateDynamicStyle -Family "paragraph" -TextProps $textProps -ParProps $parProps
                        [void]$hEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", $sName)
                    }
                    [void]$OdfTarget.AppendChild($hEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $hEl -TargetDoc $TargetDoc
                }
                "p" {
                    $pEl = $TargetDoc.CreateElement("text", "p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    if ($parProps.Count -gt 0 -or $textProps.Count -gt 0) {
                        $sName = Get-OrCreateDynamicStyle -Family "paragraph" -TextProps $textProps -ParProps $parProps
                        [void]$pEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", $sName)
                    }
                    [void]$OdfTarget.AppendChild($pEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $pEl -TargetDoc $TargetDoc
                }
                "div" {
                    $pEl = $TargetDoc.CreateElement("text", "p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$OdfTarget.AppendChild($pEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $pEl -TargetDoc $TargetDoc
                }
                "blockquote" {
                    $pEl = $TargetDoc.CreateElement("text", "p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$pEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "P_Quote")
                    [void]$OdfTarget.AppendChild($pEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $pEl -TargetDoc $TargetDoc
                }
                { $_ -in @("b", "strong") } {
                    $spanEl = $TargetDoc.CreateElement("text", "span", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$spanEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "T_Bold")
                    [void]$OdfTarget.AppendChild($spanEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $spanEl -TargetDoc $TargetDoc
                }
                { $_ -in @("i", "em") } {
                    $spanEl = $TargetDoc.CreateElement("text", "span", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$spanEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "T_Italic")
                    [void]$OdfTarget.AppendChild($spanEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $spanEl -TargetDoc $TargetDoc
                }
                "u" {
                    $spanEl = $TargetDoc.CreateElement("text", "span", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$spanEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "T_Underline")
                    [void]$OdfTarget.AppendChild($spanEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $spanEl -TargetDoc $TargetDoc
                }
                { $_ -in @("s", "strike") } {
                    $spanEl = $TargetDoc.CreateElement("text", "span", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$spanEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "T_Strike")
                    [void]$OdfTarget.AppendChild($spanEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $spanEl -TargetDoc $TargetDoc
                }
                "span" {
                    $spanEl = $TargetDoc.CreateElement("text", "span", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    if ($textProps.Count -gt 0) {
                        $sName = Get-OrCreateDynamicStyle -Family "text" -TextProps $textProps -ParProps @{}
                        [void]$spanEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", $sName)
                    }
                    [void]$OdfTarget.AppendChild($spanEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $spanEl -TargetDoc $TargetDoc
                }
                "font" {
                    $spanEl = $TargetDoc.CreateElement("text", "span", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    $fProps = @{}
                    if ($child.Attributes["color"]) { $fProps["color"] = $child.Attributes["color"].Value }
                    if ($child.Attributes["face"]) { $fProps["font-name"] = $child.Attributes["face"].Value.Split(",")[0].Trim("'", '"', ' ') }
                    if ($child.Attributes["size"]) { $fProps["font-size"] = "$([int]$child.Attributes['size'].Value * 3 + 6)pt" }
                    if ($fProps.Count -gt 0) {
                        $sName = Get-OrCreateDynamicStyle -Family "text" -TextProps $fProps -ParProps @{}
                        [void]$spanEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", $sName)
                    }
                    [void]$OdfTarget.AppendChild($spanEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $spanEl -TargetDoc $TargetDoc
                }
                "a" {
                    $aEl = $TargetDoc.CreateElement("text", "a", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    $href = if ($child.Attributes["href"]) { $child.Attributes["href"].Value } else { "#" }
                    [void]$aEl.SetAttribute("href", "http://www.w3.org/1999/xlink", $href)
                    [void]$OdfTarget.AppendChild($aEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $aEl -TargetDoc $TargetDoc
                }
                { $_ -in @("ul", "ol") } {
                    $listEl = $TargetDoc.CreateElement("text", "list", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    if ($_ -eq "ol") {
                        [void]$listEl.SetAttribute("style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0", "L_Ordered")
                    }
                    [void]$OdfTarget.AppendChild($listEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $listEl -TargetDoc $TargetDoc
                }
                "li" {
                    $itemEl = $TargetDoc.CreateElement("text", "list-item", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    $pEl = $TargetDoc.CreateElement("text", "p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$itemEl.AppendChild($pEl)
                    [void]$OdfTarget.AppendChild($itemEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $pEl -TargetDoc $TargetDoc
                }
                "table" {
                    $tblEl = $TargetDoc.CreateElement("table", "table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
                    [void]$OdfTarget.AppendChild($tblEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $tblEl -TargetDoc $TargetDoc
                }
                "tr" {
                    $trEl = $TargetDoc.CreateElement("table", "table-row", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
                    [void]$OdfTarget.AppendChild($trEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $trEl -TargetDoc $TargetDoc
                }
                { $_ -in @("td", "th") } {
                    $tdEl = $TargetDoc.CreateElement("table", "table-cell", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
                    $pEl = $TargetDoc.CreateElement("text", "p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$tdEl.AppendChild($pEl)
                    [void]$OdfTarget.AppendChild($tdEl)
                    Import-HtmlToOdf -HNode $child -OdfTarget $pEl -TargetDoc $TargetDoc
                }
                "br" {
                    $brEl = $TargetDoc.CreateElement("text", "line-break", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                    [void]$OdfTarget.AppendChild($brEl)
                }
                default {
                    Import-HtmlToOdf -HNode $child -OdfTarget $OdfTarget -TargetDoc $TargetDoc
                }
            }
        }
    }

    Import-HtmlToOdf -HNode $htmlDoc.DocumentElement -OdfTarget $bodyTextNode -TargetDoc $xml

    if ($bodyTextNode.ChildNodes.Count -eq 0) {
        $pEl = $xml.CreateElement("text", "p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        [void]$bodyTextNode.AppendChild($pEl)
    }

    $zip = [System.IO.Compression.ZipFile]::Open($OdtPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry("content.xml")
        if ($entry) { $entry.Delete() }
        $newEntry = $zip.CreateEntry("content.xml", [System.IO.Compression.CompressionLevel]::Optimal)
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($newEntry.Open(), $utf8NoBom)
        try {
            $xml.Save($writer)
        } finally {
            $writer.Close()
        }
    } finally {
        $zip.Dispose()
    }

    return $true
}

# ==============================================================================
# DOCX (OPENXML) PARSER & SERIALIZER
# ==============================================================================

function Convert-DocxRunToHtml {
    param(
        [System.Xml.XmlNode]$RunNode,
        [System.Xml.XmlNamespaceManager]$Ns
    )
    $textParts = @()
    foreach ($child in $RunNode.ChildNodes) {
        if ($child.LocalName -eq "t") {
            $textParts += [System.Net.WebUtility]::HtmlEncode($child.InnerText)
        } elseif ($child.LocalName -eq "br") {
            $textParts += "<br/>"
        } elseif ($child.LocalName -eq "tab") {
            $textParts += "&emsp;"
        }
    }
    $inner = $textParts -join ""
    if (-not $inner) { return "" }

    $rPr = $RunNode.SelectSingleNode("w:rPr", $Ns)
    if (-not $rPr) { return $inner }

    $styles = @()
    $colorNode = $rPr.SelectSingleNode("w:color", $Ns)
    if ($colorNode) {
        $val = $colorNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        if ($val -and $val -ne "auto") {
            if (-not $val.StartsWith("#")) { $val = "#$val" }
            $styles += "color: $($val.ToLower())"
        }
    }
    $shdNode = $rPr.SelectSingleNode("w:shd", $Ns)
    if ($shdNode) {
        $fill = $shdNode.GetAttribute("fill", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        if ($fill -and $fill -ne "auto") {
            if (-not $fill.StartsWith("#")) { $fill = "#$fill" }
            $styles += "background-color: $($fill.ToLower())"
        }
    }
    $hlNode = $rPr.SelectSingleNode("w:highlight", $Ns)
    if ($hlNode -and $styles -notmatch "background-color") {
        $hVal = $hlNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        if ($hVal) { $styles += "background-color: $hVal" }
    }
    $szNode = $rPr.SelectSingleNode("w:sz", $Ns)
    if ($szNode) {
        $halfPts = 0
        if ([int]::TryParse($szNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"), [ref]$halfPts) -and $halfPts -gt 0) {
            $pts = [int]($halfPts / 2)
            $styles += "font-size: ${pts}pt"
        }
    }
    $fNode = $rPr.SelectSingleNode("w:rFonts", $Ns)
    if ($fNode) {
        $ascii = $fNode.GetAttribute("ascii", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        if ($ascii) { $styles += "font-family: '$ascii', sans-serif" }
    }

    if ($styles.Count -gt 0) {
        $inner = "<span style=""$($styles -join '; ')"">$inner</span>"
    }
    if ($rPr.SelectSingleNode("w:strike", $Ns)) { $inner = "<s>$inner</s>" }
    if ($rPr.SelectSingleNode("w:u", $Ns)) { $inner = "<u>$inner</u>" }
    if ($rPr.SelectSingleNode("w:i", $Ns)) { $inner = "<i>$inner</i>" }
    if ($rPr.SelectSingleNode("w:b", $Ns)) { $inner = "<b>$inner</b>" }

    return $inner
}

function Convert-DocxParagraphChildrenToHtml {
    param(
        [System.Xml.XmlNode]$PNode,
        [System.Xml.XmlNamespaceManager]$Ns,
        [hashtable]$RelsMap
    )
    $parts = @()
    foreach ($child in $PNode.ChildNodes) {
        if ($child.LocalName -eq "r") {
            $parts += Convert-DocxRunToHtml -RunNode $child -Ns $Ns
        } elseif ($child.LocalName -eq "hyperlink") {
            $rId = $child.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            $href = if ($rId -and $RelsMap.ContainsKey($rId)) { $RelsMap[$rId] } else { "#" }
            $linkInner = ""
            foreach ($hc in $child.ChildNodes) {
                if ($hc.LocalName -eq "r") {
                    $linkInner += Convert-DocxRunToHtml -RunNode $hc -Ns $Ns
                }
            }
            $parts += "<a href=""$href"">$linkInner</a>"
        }
    }
    return ($parts -join "")
}

function Convert-DocxToHtml {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DocxPath
    )
    if (-not (Test-Path $DocxPath)) { throw "File not found: $DocxPath" }

    $docXmlStr = Get-ZipEntryText -ZipPath $DocxPath -EntryName "word/document.xml"
    if (-not $docXmlStr) { throw "Missing word/document.xml in DOCX file: $DocxPath" }

    $relsMap = @{}
    $relsXmlStr = Get-ZipEntryText -ZipPath $DocxPath -EntryName "word/_rels/document.xml.rels"
    if ($relsXmlStr) {
        try {
            [xml]$relsXml = $relsXmlStr
            foreach ($rel in $relsXml.DocumentElement.ChildNodes) {
                if ($rel.LocalName -eq "Relationship") {
                    $id = $rel.GetAttribute("Id")
                    $target = $rel.GetAttribute("Target")
                    if ($id -and $target) { $relsMap[$id] = $target }
                }
            }
        } catch {}
    }

    [xml]$xml = $docXmlStr
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $bodyNode = $xml.SelectSingleNode("//w:body", $ns)
    if (-not $bodyNode) {
        return @{
            title = [System.IO.Path]::GetFileNameWithoutExtension($DocxPath)
            format = "docx"
            html = "<p></p>"
            styles = ""
        }
    }

    $htmlParts = @()
    $currentListType = $null
    $listBuffer = @()

    function Flush-ListBuffer {
        if ($script:currentListType -and $script:listBuffer.Count -gt 0) {
            $script:htmlParts += "<$($script:currentListType)>$($script:listBuffer -join '')</$($script:currentListType)>"
            $script:listBuffer = @()
            $script:currentListType = $null
        }
    }

    $script:htmlParts = @()
    $script:currentListType = $null
    $script:listBuffer = @()

    foreach ($child in $bodyNode.ChildNodes) {
        if ($child.LocalName -eq "p") {
            $pPr = $child.SelectSingleNode("w:pPr", $ns)
            $pStyle = ""
            $align = ""
            $numId = ""

            if ($pPr) {
                $pStyleNode = $pPr.SelectSingleNode("w:pStyle", $ns)
                if ($pStyleNode) {
                    $pStyle = $pStyleNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                }
                $jcNode = $pPr.SelectSingleNode("w:jc", $ns)
                if ($jcNode) {
                    $jcVal = $jcNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                    if ($jcVal -eq "both") { $align = "justify" }
                    elseif ($jcVal -in @("left", "center", "right")) { $align = $jcVal }
                }
                $numNode = $pPr.SelectSingleNode("w:numPr/w:numId", $ns)
                if ($numNode) {
                    $numId = $numNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                }
            }

            $innerHtml = Convert-DocxParagraphChildrenToHtml -PNode $child -Ns $ns -RelsMap $relsMap

            if ($numId) {
                $targetListType = if ($numId -eq "2") { "ol" } else { "ul" }
                if ($script:currentListType -and $script:currentListType -ne $targetListType) {
                    Flush-ListBuffer
                }
                $script:currentListType = $targetListType
                $script:listBuffer += "<li>$innerHtml</li>"
                continue
            } else {
                Flush-ListBuffer
            }

            $styleAttr = if ($align) { " style=""text-align: $align""" } else { "" }

            if ($pStyle -match "^(?i)heading\s*([1-6])$") {
                $lvl = $matches[1]
                $script:htmlParts += "<h$lvl$styleAttr>$innerHtml</h$lvl>"
            } elseif ($pStyle -match "^(?i)quote|intensequote|blockquote$") {
                $script:htmlParts += "<blockquote$styleAttr>$innerHtml</blockquote>"
            } else {
                if ([string]::IsNullOrWhiteSpace($innerHtml)) { $innerHtml = "<br/>" }
                $script:htmlParts += "<p$styleAttr>$innerHtml</p>"
            }
        }
        elseif ($child.LocalName -eq "tbl") {
            Flush-ListBuffer
            $rowsHtml = @()
            foreach ($tr in $child.SelectNodes("w:tr", $ns)) {
                $cellsHtml = @()
                foreach ($tc in $tr.SelectNodes("w:tc", $ns)) {
                    $cellContent = @()
                    foreach ($cp in $tc.SelectNodes("w:p", $ns)) {
                        $cellContent += Convert-DocxParagraphChildrenToHtml -PNode $cp -Ns $ns -RelsMap $relsMap
                    }
                    $cellsHtml += "<td>$($cellContent -join '<br/>')</td>"
                }
                $rowsHtml += "<tr>$($cellsHtml -join '')</tr>"
            }
            $script:htmlParts += "<table class=""doc-table"">$($rowsHtml -join '')</table>"
        }
    }
    Flush-ListBuffer

    return @{
        title = [System.IO.Path]::GetFileNameWithoutExtension($DocxPath)
        format = "docx"
        html = ($script:htmlParts -join "`n")
        styles = ""
    }
}

function Convert-HtmlToDocx {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DocxPath,
        [Parameter(Mandatory=$true)]
        [string]$HtmlContent
    )

    if (-not (Test-Path $DocxPath)) {
        [void](New-BlankDocx -DocxPath $DocxPath)
    }

    $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    $rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    $existingXmlStr = Get-ZipEntryText -ZipPath $DocxPath -EntryName "word/document.xml"
    [xml]$docXml = $existingXmlStr
    $ns = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
    $ns.AddNamespace("w", $wNs)
    $ns.AddNamespace("r", $rNs)

    $bodyNode = $docXml.SelectSingleNode("//w:body", $ns)
    if (-not $bodyNode) {
        $bodyNode = $docXml.CreateElement("w", "body", $wNs)
        [void]$docXml.DocumentElement.AppendChild($bodyNode)
    }

    # Preserve existing w:sectPr at the end of w:body
    $sectPrNode = $bodyNode.SelectSingleNode("w:sectPr", $ns)
    if ($sectPrNode) {
        $sectPrClone = $sectPrNode.CloneNode($true)
    } else {
        $sectPrClone = $docXml.CreateElement("w", "sectPr", $wNs)
        $pgSz = $docXml.CreateElement("w", "pgSz", $wNs)
        [void]$pgSz.SetAttribute("w", $wNs, "12240")
        [void]$pgSz.SetAttribute("h", $wNs, "15840")
        [void]$sectPrClone.AppendChild($pgSz)
    }

    [void]$bodyNode.RemoveAll()

    # Load or initialize word/_rels/document.xml.rels
    $relsStr = Get-ZipEntryText -ZipPath $DocxPath -EntryName "word/_rels/document.xml.rels"
    if (-not $relsStr) {
        $relsStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
    }
    [xml]$relsXml = $relsStr
    $relNsUri = "http://schemas.openxmlformats.org/package/2006/relationships"

    # Ensure numbering relationship exists
    $hasNumRel = $false
    $maxRId = 10
    foreach ($rel in $relsXml.DocumentElement.ChildNodes) {
        if ($rel.LocalName -eq "Relationship") {
            if ($rel.GetAttribute("Target") -eq "numbering.xml") { $hasNumRel = $true }
            $idStr = $rel.GetAttribute("Id")
            if ($idStr -match "^rId(\d+)$") {
                $num = [int]$matches[1]
                if ($num -gt $maxRId) { $maxRId = $num }
            }
        }
    }
    if (-not $hasNumRel) {
        $maxRId++
        $numRel = $relsXml.CreateElement("Relationship", $relNsUri)
        [void]$numRel.SetAttribute("Id", "rId$maxRId")
        [void]$numRel.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering")
        [void]$numRel.SetAttribute("Target", "numbering.xml")
        [void]$relsXml.DocumentElement.AppendChild($numRel)
    }

    $script:nextRelId = $maxRId + 1
    function Register-DocxHyperlink {
        param([string]$Url)
        $rId = "rId$($script:nextRelId)"
        $script:nextRelId++
        $relEl = $relsXml.CreateElement("Relationship", $relNsUri)
        [void]$relEl.SetAttribute("Id", $rId)
        [void]$relEl.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
        [void]$relEl.SetAttribute("Target", $Url)
        [void]$relEl.SetAttribute("TargetMode", "External")
        [void]$relsXml.DocumentElement.AppendChild($relEl)
        return $rId
    }

    $normalizedHtml = Normalize-HtmlForXml -html $HtmlContent
    $wrappedHtml = "<root>" + $normalizedHtml + "</root>"
    $htmlDoc = New-Object System.Xml.XmlDocument
    try {
        $htmlDoc.LoadXml($wrappedHtml)
    } catch {
        $safeText = [System.Net.WebUtility]::HtmlEncode($HtmlContent)
        $htmlDoc.LoadXml("<root><p>$safeText</p></root>")
    }

    function Parse-InlineCss {
        param([System.Xml.XmlNode]$Node)
        $props = @{}
        if ($Node.Attributes -and $Node.Attributes["style"]) {
            foreach ($pair in $Node.Attributes["style"].Value.Split(";")) {
                $kv = $pair.Split(":")
                if ($kv.Length -eq 2) {
                    $k = $kv[0].Trim().ToLower()
                    $v = $kv[1].Trim().Trim("'", '"')
                    $props[$k] = $v
                }
            }
        }
        return $props
    }

    function Add-DocxRunsFromHtml {
        param(
            [System.Xml.XmlNode]$HNode,
            [System.Xml.XmlNode]$ParentWNode,
            [hashtable]$FormatState
        )
        foreach ($child in $HNode.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Text) {
                $txt = $child.Value
                if ($txt.Length -eq 0) { continue }
                $rEl = $docXml.CreateElement("w", "r", $wNs)
                $rPr = $docXml.CreateElement("w", "rPr", $wNs)
                $hasRPr = $false

                if ($FormatState["bold"]) {
                    [void]$rPr.AppendChild($docXml.CreateElement("w", "b", $wNs))
                    $hasRPr = $true
                }
                if ($FormatState["italic"]) {
                    [void]$rPr.AppendChild($docXml.CreateElement("w", "i", $wNs))
                    $hasRPr = $true
                }
                if ($FormatState["underline"]) {
                    $uEl = $docXml.CreateElement("w", "u", $wNs)
                    [void]$uEl.SetAttribute("val", $wNs, "single")
                    [void]$rPr.AppendChild($uEl)
                    $hasRPr = $true
                }
                if ($FormatState["strike"]) {
                    [void]$rPr.AppendChild($docXml.CreateElement("w", "strike", $wNs))
                    $hasRPr = $true
                }
                if ($FormatState["color"]) {
                    $cVal = $FormatState["color"].TrimStart("#")
                    if ($cVal.Length -eq 6) {
                        $cEl = $docXml.CreateElement("w", "color", $wNs)
                        [void]$cEl.SetAttribute("val", $wNs, $cVal.ToUpper())
                        [void]$rPr.AppendChild($cEl)
                        $hasRPr = $true
                    }
                }
                if ($FormatState["bg"]) {
                    $bgVal = $FormatState["bg"].TrimStart("#")
                    if ($bgVal.Length -eq 6) {
                        $shdEl = $docXml.CreateElement("w", "shd", $wNs)
                        [void]$shdEl.SetAttribute("val", $wNs, "clear")
                        [void]$shdEl.SetAttribute("color", $wNs, "auto")
                        [void]$shdEl.SetAttribute("fill", $wNs, $bgVal.ToUpper())
                        [void]$rPr.AppendChild($shdEl)
                        $hasRPr = $true
                    }
                }
                if ($FormatState["font"]) {
                    $fontName = $FormatState["font"].Split(",")[0].Trim("'", '"', ' ')
                    $rfEl = $docXml.CreateElement("w", "rFonts", $wNs)
                    [void]$rfEl.SetAttribute("ascii", $wNs, $fontName)
                    [void]$rfEl.SetAttribute("hAnsi", $wNs, $fontName)
                    [void]$rPr.AppendChild($rfEl)
                    $hasRPr = $true
                }
                if ($FormatState["size"]) {
                    $szStr = $FormatState["size"] -replace "[^0-9.]", ""
                    $pts = 0
                    if ([double]::TryParse($szStr, [ref]$pts) -and $pts -gt 0) {
                        $halfPts = [int]($pts * 2)
                        $szEl = $docXml.CreateElement("w", "sz", $wNs)
                        [void]$szEl.SetAttribute("val", $wNs, "$halfPts")
                        [void]$rPr.AppendChild($szEl)
                        $hasRPr = $true
                    }
                }

                if ($hasRPr) { [void]$rEl.AppendChild($rPr) }
                $tEl = $docXml.CreateElement("w", "t", $wNs)
                [void]$tEl.SetAttribute("xml:space", "preserve")
                $tEl.InnerText = $txt
                [void]$rEl.AppendChild($tEl)
                [void]$ParentWNode.AppendChild($rEl)
                continue
            }

            $tag = $child.LocalName.ToLower()
            if ($tag -eq "br") {
                $rEl = $docXml.CreateElement("w", "r", $wNs)
                [void]$rEl.AppendChild($docXml.CreateElement("w", "br", $wNs))
                [void]$ParentWNode.AppendChild($rEl)
                continue
            }

            $nextState = $FormatState.Clone()
            $css = Parse-InlineCss -Node $child
            if ($css.ContainsKey("color")) { $nextState["color"] = $css["color"] }
            if ($css.ContainsKey("background-color")) { $nextState["bg"] = $css["background-color"] }
            if ($css.ContainsKey("font-family")) { $nextState["font"] = $css["font-family"] }
            if ($css.ContainsKey("font-size")) { $nextState["size"] = $css["font-size"] }
            if ($css.ContainsKey("font-weight") -and $css["font-weight"] -eq "bold") { $nextState["bold"] = $true }
            if ($css.ContainsKey("font-style") -and $css["font-style"] -eq "italic") { $nextState["italic"] = $true }
            if ($css.ContainsKey("text-decoration")) {
                if ($css["text-decoration"] -match "underline") { $nextState["underline"] = $true }
                if ($css["text-decoration"] -match "line-through") { $nextState["strike"] = $true }
            }

            if ($tag -in @("b", "strong")) { $nextState["bold"] = $true }
            elseif ($tag -in @("i", "em")) { $nextState["italic"] = $true }
            elseif ($tag -eq "u") { $nextState["underline"] = $true }
            elseif ($tag -in @("s", "strike")) { $nextState["strike"] = $true }
            elseif ($tag -eq "font") {
                if ($child.Attributes["color"]) { $nextState["color"] = $child.Attributes["color"].Value }
                if ($child.Attributes["face"]) { $nextState["font"] = $child.Attributes["face"].Value }
                if ($child.Attributes["size"]) { $nextState["size"] = "$([int]$child.Attributes['size'].Value * 3 + 6)pt" }
            }
            elseif ($tag -eq "a") {
                $href = if ($child.Attributes["href"]) { $child.Attributes["href"].Value } else { "https://example.com" }
                $rId = Register-DocxHyperlink -Url $href
                $hlEl = $docXml.CreateElement("w", "hyperlink", $wNs)
                [void]$hlEl.SetAttribute("id", $rNs, $rId)
                $nextState["underline"] = $true
                if (-not $nextState["color"]) { $nextState["color"] = "#2563eb" }
                Add-DocxRunsFromHtml -HNode $child -ParentWNode $hlEl -FormatState $nextState
                [void]$ParentWNode.AppendChild($hlEl)
                continue
            }

            Add-DocxRunsFromHtml -HNode $child -ParentWNode $ParentWNode -FormatState $nextState
        }
    }

    function Add-DocxParagraph {
        param(
            [System.Xml.XmlNode]$HNode,
            [string]$StyleId = "",
            [string]$NumId = ""
        )
        $pEl = $docXml.CreateElement("w", "p", $wNs)
        $pPr = $docXml.CreateElement("w", "pPr", $wNs)
        $hasPPr = $false

        if ($StyleId) {
            $psEl = $docXml.CreateElement("w", "pStyle", $wNs)
            [void]$psEl.SetAttribute("val", $wNs, $StyleId)
            [void]$pPr.AppendChild($psEl)
            $hasPPr = $true
        }
        if ($NumId) {
            $numPr = $docXml.CreateElement("w", "numPr", $wNs)
            $ilvl = $docXml.CreateElement("w", "ilvl", $wNs)
            [void]$ilvl.SetAttribute("val", $wNs, "0")
            $nId = $docXml.CreateElement("w", "numId", $wNs)
            [void]$nId.SetAttribute("val", $wNs, $NumId)
            [void]$numPr.AppendChild($ilvl)
            [void]$numPr.AppendChild($nId)
            [void]$pPr.AppendChild($numPr)
            $hasPPr = $true
        }

        $css = Parse-InlineCss -Node $HNode
        if ($css.ContainsKey("text-align")) {
            $align = $css["text-align"].ToLower()
            $jcVal = if ($align -eq "justify") { "both" } else { $align }
            if ($jcVal -in @("left", "center", "right", "both")) {
                $jcEl = $docXml.CreateElement("w", "jc", $wNs)
                [void]$jcEl.SetAttribute("val", $wNs, $jcVal)
                [void]$pPr.AppendChild($jcEl)
                $hasPPr = $true
            }
        }

        if ($hasPPr) { [void]$pEl.AppendChild($pPr) }

        $initFormat = @{}
        if ($css.ContainsKey("color")) { $initFormat["color"] = $css["color"] }
        if ($css.ContainsKey("font-size")) { $initFormat["size"] = $css["font-size"] }
        if ($css.ContainsKey("font-family")) { $initFormat["font"] = $css["font-family"] }

        Add-DocxRunsFromHtml -HNode $HNode -ParentWNode $pEl -FormatState $initFormat
        [void]$bodyNode.AppendChild($pEl)
    }

    function Import-HtmlBlockToDocx {
        param([System.Xml.XmlNode]$Container)
        foreach ($child in $Container.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Text) {
                if (-not [string]::IsNullOrWhiteSpace($child.Value)) {
                    $pEl = $docXml.CreateElement("w", "p", $wNs)
                    $rEl = $docXml.CreateElement("w", "r", $wNs)
                    $tEl = $docXml.CreateElement("w", "t", $wNs)
                    $tEl.InnerText = $child.Value
                    [void]$rEl.AppendChild($tEl)
                    [void]$pEl.AppendChild($rEl)
                    [void]$bodyNode.AppendChild($pEl)
                }
                continue
            }

            $tag = $child.LocalName.ToLower()
            switch ($tag) {
                { $_ -match "^h([1-6])$" } {
                    $lvl = $_.Substring(1)
                    if ([int]$lvl -gt 3) { $lvl = "3" }
                    Add-DocxParagraph -HNode $child -StyleId "Heading$lvl"
                }
                "p" { Add-DocxParagraph -HNode $child }
                "div" { Add-DocxParagraph -HNode $child }
                "blockquote" { Add-DocxParagraph -HNode $child -StyleId "Quote" }
                "ul" {
                    foreach ($li in $child.ChildNodes) {
                        if ($li.LocalName -eq "li") {
                            Add-DocxParagraph -HNode $li -NumId "1"
                        }
                    }
                }
                "ol" {
                    foreach ($li in $child.ChildNodes) {
                        if ($li.LocalName -eq "li") {
                            Add-DocxParagraph -HNode $li -NumId "2"
                        }
                    }
                }
                "table" {
                    $tblEl = $docXml.CreateElement("w", "tbl", $wNs)
                    $tblPr = $docXml.CreateElement("w", "tblPr", $wNs)
                    $tblW = $docXml.CreateElement("w", "tblW", $wNs)
                    [void]$tblW.SetAttribute("w", $wNs, "5000")
                    [void]$tblW.SetAttribute("type", $wNs, "pct")
                    [void]$tblPr.AppendChild($tblW)
                    [void]$tblEl.AppendChild($tblPr)

                    $rows = $child.SelectNodes(".//tr")
                    foreach ($tr in $rows) {
                        $trEl = $docXml.CreateElement("w", "tr", $wNs)
                        foreach ($cell in $tr.ChildNodes) {
                            if ($cell.LocalName -in @("td", "th")) {
                                $tcEl = $docXml.CreateElement("w", "tc", $wNs)
                                $pEl = $docXml.CreateElement("w", "p", $wNs)
                                $cellFmt = @{}
                                if ($cell.LocalName -eq "th") { $cellFmt["bold"] = $true }
                                Add-DocxRunsFromHtml -HNode $cell -ParentWNode $pEl -FormatState $cellFmt
                                [void]$tcEl.AppendChild($pEl)
                                [void]$trEl.AppendChild($tcEl)
                            }
                        }
                        [void]$tblEl.AppendChild($trEl)
                    }
                    [void]$bodyNode.AppendChild($tblEl)
                }
                default {
                    Import-HtmlBlockToDocx -Container $child
                }
            }
        }
    }

    Import-HtmlBlockToDocx -Container $htmlDoc.DocumentElement

    # Ensure at least one paragraph precedes sectPr
    if ($bodyNode.ChildNodes.Count -eq 0) {
        [void]$bodyNode.AppendChild($docXml.CreateElement("w", "p", $wNs))
    }

    # Re-append preserved w:sectPr at the end of w:body
    [void]$bodyNode.AppendChild($sectPrClone)

    # Save updated entries to the DOCX zip package
    $zip = [System.IO.Compression.ZipFile]::Open($DocxPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        # Ensure numbering.xml exists
        if (-not $zip.GetEntry("word/numbering.xml")) {
            Set-ZipEntryText -Zip $zip -EntryName "word/numbering.xml" -Content (Get-DefaultDocxNumberingXml)
        }
        # Ensure styles.xml exists
        if (-not $zip.GetEntry("word/styles.xml")) {
            Set-ZipEntryText -Zip $zip -EntryName "word/styles.xml" -Content (Get-DefaultDocxStylesXml)
        }

        # Ensure [Content_Types].xml declares numbering.xml
        $ctEntry = $zip.GetEntry("[Content_Types].xml")
        if ($ctEntry) {
            $r = New-Object System.IO.StreamReader($ctEntry.Open(), [System.Text.Encoding]::UTF8)
            $ctText = $r.ReadToEnd()
            $r.Close()
            if ($ctText -notmatch "numbering\.xml") {
                [xml]$ctXml = $ctText
                $ov = $ctXml.CreateElement("Override", $ctXml.DocumentElement.NamespaceURI)
                [void]$ov.SetAttribute("PartName", "/word/numbering.xml")
                [void]$ov.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
                [void]$ctXml.DocumentElement.AppendChild($ov)
                $sw = New-Object System.IO.StringWriter
                $ctXml.Save($sw)
                Set-ZipEntryText -Zip $zip -EntryName "[Content_Types].xml" -Content $sw.ToString()
            }
        }

        $swDoc = New-Object System.IO.StringWriter
        $docXml.Save($swDoc)
        Set-ZipEntryText -Zip $zip -EntryName "word/document.xml" -Content $swDoc.ToString()

        $swRels = New-Object System.IO.StringWriter
        $relsXml.Save($swRels)
        Set-ZipEntryText -Zip $zip -EntryName "word/_rels/document.xml.rels" -Content $swRels.ToString()
    } finally {
        $zip.Dispose()
    }

    return $true
}

# ==============================================================================
# UNIFIED DISPATCHERS & CROSS-FORMAT CONVERSION
# ==============================================================================

function Convert-DocumentToHtml {
    param(
        [Parameter(Mandatory=$true)]
        [Alias("Path")]
        [string]$DocPath
    )
    $ext = [System.IO.Path]::GetExtension($DocPath).ToLower()
    if ($ext -eq ".docx") {
        return (Convert-DocxToHtml -DocxPath $DocPath)
    } elseif ($ext -eq ".odt") {
        return (Convert-OdtToHtml -OdtPath $DocPath)
    } else {
        throw "Unsupported document extension: $ext (expected .odt or .docx)"
    }
}

function Convert-HtmlToDocument {
    param(
        [Parameter(Mandatory=$true)]
        [Alias("Path")]
        [string]$DocPath,
        [Parameter(Mandatory=$true)]
        [string]$HtmlContent
    )
    $ext = [System.IO.Path]::GetExtension($DocPath).ToLower()
    if ($ext -eq ".docx") {
        return (Convert-HtmlToDocx -DocxPath $DocPath -HtmlContent $HtmlContent)
    } elseif ($ext -eq ".odt") {
        return (Convert-HtmlToOdt -OdtPath $DocPath -HtmlContent $HtmlContent)
    } else {
        throw "Unsupported document extension: $ext (expected .odt or .docx)"
    }
}

function Convert-HtmlToMarkdown {
    param([string]$HtmlContent)
    $md = $HtmlContent
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<h1[^>]*>(.*?)</h1>", "# `$1`n`n")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<h2[^>]*>(.*?)</h2>", "## `$1`n`n")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<h3[^>]*>(.*?)</h3>", "### `$1`n`n")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<blockquote[^>]*>(.*?)</blockquote>", "> `$1`n`n")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<a[^>]*href=['""]([^'""]+)['""][^>]*>(.*?)</a>", "[`$2](`$1)")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<(b|strong)[^>]*>(.*?)</\1>", "**`$2**")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<(i|em)[^>]*>(.*?)</\1>", "*`$2*")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<(s|strike)[^>]*>(.*?)</\1>", "~~`$2~~")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<li[^>]*>(.*?)</li>", "- `$1`n")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)</?(ul|ol)[^>]*>", "`n")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<p[^>]*>(.*?)</p>", "`$1`n`n")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "(?is)<br\s*/?>", "`n")
    $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<[^>]+>", "")
    $decoded = [System.Net.WebUtility]::HtmlDecode($md)
    $decoded = [System.Text.RegularExpressions.Regex]::Replace($decoded, "(\r?\n){3,}", "`n`n")
    return $decoded.Trim()
}

function Convert-DocumentFormat {
    param(
        [string]$SourcePath = "",
        [Parameter(Mandatory=$true)]
        [string]$TargetPath,
        [string]$HtmlContent = ""
    )

    if (-not $HtmlContent) {
        if (-not $SourcePath -or -not (Test-Path $SourcePath)) {
            throw "Either HtmlContent or a valid SourcePath must be provided"
        }
        $parsed = Convert-DocumentToHtml -DocPath $SourcePath
        $HtmlContent = $parsed.html
    }

    $targetExt = [System.IO.Path]::GetExtension($TargetPath).ToLower()
    $parentDir = Split-Path -Parent $TargetPath
    if ($parentDir -and -not (Test-Path $parentDir)) {
        [void](New-Item -ItemType Directory -Path $parentDir -Force)
    }

    if ($targetExt -eq ".docx") {
        if (-not (Test-Path $TargetPath)) {
            [void](New-BlankDocx -DocxPath $TargetPath)
        }
        return (Convert-HtmlToDocx -DocxPath $TargetPath -HtmlContent $HtmlContent)
    }
    elseif ($targetExt -eq ".odt") {
        if (-not (Test-Path $TargetPath)) {
            [void](New-BlankOdt -OdtPath $TargetPath)
        }
        return (Convert-HtmlToOdt -OdtPath $TargetPath -HtmlContent $HtmlContent)
    }
    elseif ($targetExt -in @(".md", ".markdown")) {
        $mdText = Convert-HtmlToMarkdown -HtmlContent $HtmlContent
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($TargetPath, $mdText, $utf8NoBom)
        return $true
    }
    else {
        throw "Unsupported target conversion format: $targetExt"
    }
}

# Standalone CLI execution
if ($Read) {
    if (-not (Test-Path $Read)) {
        Write-Error "File not found: $Read"
        exit 1
    }
    $doc = Convert-DocumentToHtml -DocPath $Read
    if ($Format -eq "html") {
        Write-Output $doc.html
    } elseif ($Format -eq "text") {
        $text = [System.Text.RegularExpressions.Regex]::Replace($doc.html, "<[^>]+>", " ")
        $decoded = [System.Net.WebUtility]::HtmlDecode($text)
        Write-Output $decoded.Trim()
    } else {
        Write-Output (Convert-HtmlToMarkdown -HtmlContent $doc.html)
    }
}
