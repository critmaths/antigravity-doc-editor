<#
.SYNOPSIS
    ODT Document Converter for Antigravity ODT Editor.
    Provides bidirectional conversion between OpenDocument Text (.odt) and HTML/Markdown.
#>

param(
    [string]$Read = "",
    [ValidateSet("html", "markdown", "text")]
    [string]$Format = "markdown"
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-OdtZipEntryText {
    param(
        [string]$OdtPath,
        [string]$EntryName
    )
    $zip = [System.IO.Compression.ZipFile]::OpenRead($OdtPath)
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
                if (-not $level -or $level -gt 4) { $level = 1 }
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
                $styleAttr = if ($StylesMap.ContainsKey($styleName)) { " style=""$($StylesMap[$styleName])""" } else { "" }
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                }
                return "<span$styleAttr>$innerHtml</span>"
            }
            "list" {
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                }
                return "<ul>$innerHtml</ul>"
            }
            "list-item" {
                $innerHtml = ""
                foreach ($child in $Node.ChildNodes) {
                    $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
                }
                return "<li>$innerHtml</li>"
            }
            "line-break" {
                return "<br/>"
            }
            "tab" {
                return "&emsp;"
            }
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
        return "<table class=""odt-table"">$innerHtml</table>"
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
            $innerHtml += Convert-OdfNodeToHtml -Node $child -StylesMap $StylesMap -Ns $Ns
        }
        return "<td>$innerHtml</td>"
    }

    # Fallback: process children
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

    if (-not (Test-Path $OdtPath)) {
        throw "File not found: $OdtPath"
    }

    $contentXmlStr = Get-OdtZipEntryText -OdtPath $OdtPath -EntryName "content.xml"
    if (-not $contentXmlStr) {
        throw "Missing content.xml in ODT file: $OdtPath"
    }

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
        html = ($htmlParts -join "`n")
        styles = ""
    }
}

function Normalize-HtmlForXml {
    param([string]$html)
    
    # 1. Normalize linebreaks, hr, img
    $clean = $html.Replace("&nbsp;", " ").Replace("&emsp;", "    ")
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "<br\s*>", "<br/>")
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "<hr\s*>", "<hr/>")
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "<img([^>]+)>", "<img`$1/>")

    # 2. Decode standard HTML entities (like &mdash;, &copy;, etc.) to UTF-8, preserving XML entities
    $entityCallback = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $ent = $m.Groups[1].Value.ToLower()
        if ($ent -in @("amp", "lt", "gt", "quot", "apos")) {
            return $m.Value
        }
        $decoded = [System.Net.WebUtility]::HtmlDecode($m.Value)
        if ($decoded -eq "&") { return "&amp;" }
        if ($decoded -eq "<") { return "&lt;" }
        if ($decoded -eq ">") { return "&gt;" }
        if ($decoded -eq '"') { return "&quot;" }
        if ($decoded -eq "'") { return "&apos;" }
        return $decoded
    }
    $clean = [System.Text.RegularExpressions.Regex]::Replace($clean, "&([a-zA-Z0-9]+);", $entityCallback)

    # 3. Escape bare '&' not part of valid entity
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
        throw "Target ODT file not found: $OdtPath"
    }

    # Read existing content.xml to preserve namespaces & base structure
    $existingXmlStr = Get-OdtZipEntryText -OdtPath $OdtPath -EntryName "content.xml"
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

    # Helper styles
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

    # Normalize HTML for XML parsing
    $normalizedHtml = Normalize-HtmlForXml -html $HtmlContent
    $wrappedHtml = "<root>" + $normalizedHtml + "</root>"

    $htmlDoc = New-Object System.Xml.XmlDocument
    try {
        $htmlDoc.LoadXml($wrappedHtml)
    } catch {
        # Fallback if unrecoverable XML error
        $safeText = [System.Net.WebUtility]::HtmlEncode($HtmlContent)
        $htmlDoc.LoadXml("<root><p>$safeText</p></root>")
    }

    # Style generator for inline style attributes
    $script:dynStyleId = 100
    function Get-OrCreateDynamicStyle {
        param(
            [string]$Family,
            [hashtable]$TextProps,
            [hashtable]$ParProps
        )
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

    # Recursive import from HTML DOM to ODF XML
    function Import-HtmlToOdf {
        param(
            [System.Xml.XmlNode]$HNode,
            [System.Xml.XmlNode]$OdfTarget,
            [System.Xml.XmlDocument]$TargetDoc
        )

        foreach ($child in $HNode.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Text) {
                $textNode = $TargetDoc.CreateTextNode($child.Value)
                [void]$OdfTarget.AppendChild($textNode)
                continue
            }

            $tag = $child.LocalName.ToLower()

            # Extract inline style properties if present
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
                    if ($child.Attributes["face"]) { $fProps["font-name"] = $child.Attributes["face"].Value }
                    if ($child.Attributes["size"]) { $fProps["font-size"] = "$($child.Attributes['size'].Value * 3 + 6)pt" }
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

    # Write updated content.xml back to the zip archive
    $zip = [System.IO.Compression.ZipFile]::Open($OdtPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry("content.xml")
        if ($entry) { $entry.Delete() }
        $newEntry = $zip.CreateEntry("content.xml", [System.IO.Compression.CompressionLevel]::Optimal)
        $writer = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
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

# Standalone CLI execution
if ($Read) {
    if (-not (Test-Path $Read)) {
        Write-Error "File not found: $Read"
        exit 1
    }
    $doc = Convert-OdtToHtml -OdtPath $Read
    if ($Format -eq "html") {
        Write-Output $doc.html
    } elseif ($Format -eq "text") {
        $text = [System.Text.RegularExpressions.Regex]::Replace($doc.html, "<[^>]+>", " ")
        $decoded = [System.Net.WebUtility]::HtmlDecode($text)
        Write-Output $decoded.Trim()
    } else {
        # Markdown conversion
        $md = $doc.html
        $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<h1[^>]*>(.*?)</h1>", "# `$1`n`n")
        $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<h2[^>]*>(.*?)</h2>", "## `$1`n`n")
        $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<h3[^>]*>(.*?)</h3>", "### `$1`n`n")
        $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<p[^>]*>(.*?)</p>", "`$1`n`n")
        $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<li[^>]*>(.*?)</li>", "- `$1`n")
        $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<(b|strong)[^>]*>(.*?)</\1>", "**`$2**")
        $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<(i|em)[^>]*>(.*?)</\1>", "*`$2*")
        $md = [System.Text.RegularExpressions.Regex]::Replace($md, "<[^>]+>", "")
        $decoded = [System.Net.WebUtility]::HtmlDecode($md)
        Write-Output $decoded.Trim()
    }
}
