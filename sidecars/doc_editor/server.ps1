<#
.SYNOPSIS
    Antigravity Document Editor Sidecar HTTP Server.
    Provides local REST endpoints and serves the visual WYSIWYG editor webview.
    Supports both OpenDocument Text (.odt) and Microsoft Word OpenXML (.docx).
#>

param(
    [int]$Port = 0
)

$ErrorActionPreference = "Stop"

if ($env:ANTIGRAVITY_SIDECAR_WEB_PORT) {
    try { $Port = [int]$env:ANTIGRAVITY_SIDECAR_WEB_PORT } catch {}
}
if (-not $Port -or $Port -eq 0) {
    $Port = 58912
}

# Dot-source the unified converter
$ConverterScript = Join-Path $PSScriptRoot "doc_converter.ps1"
if (-not (Test-Path $ConverterScript)) {
    throw "Required doc_converter.ps1 not found in $PSScriptRoot"
}
. $ConverterScript

$listener = New-Object System.Net.HttpListener
$prefix = "http://127.0.0.1:$Port/"
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
} catch {
    Write-Host "Antigravity Document Editor server already listening on $prefix or port in use."
    exit 0
}

Write-Host "Antigravity Document Editor server started on $prefix"

function Send-JsonResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [object]$Data,
        [int]$StatusCode = 200
    )
    $json = $Data | ConvertTo-Json -Depth 5 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = "application/json; charset=utf-8"
    $Response.ContentLength64 = $bytes.Length
    $Response.AddHeader("Access-Control-Allow-Origin", "*")
    $Response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    $Response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Accept")
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Send-FileResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [string]$FilePath,
        [string]$ContentType = "text/html; charset=utf-8"
    )
    if (-not (Test-Path $FilePath)) {
        Send-JsonResponse -Response $Response -Data @{ error = "File not found" } -StatusCode 404
        return
    }
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    $Response.StatusCode = 200
    $Response.ContentType = $ContentType
    $Response.ContentLength64 = $bytes.Length
    $Response.AddHeader("Access-Control-Allow-Origin", "*")
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

$running = $true

while ($running -and $listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $urlPath = $request.Url.AbsolutePath
        $method = $request.HttpMethod

        # CORS preflight
        if ($method -eq "OPTIONS") {
            $response.StatusCode = 200
            $response.AddHeader("Access-Control-Allow-Origin", "*")
            $response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            $response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Accept")
            $response.OutputStream.Close()
            continue
        }

        # Shutdown endpoint (useful for tests and clean stops)
        if ($urlPath -eq "/api/shutdown") {
            Send-JsonResponse -Response $response -Data @{ status = "shutting_down" }
            $running = $false
            break
        }

        # Static assets
        if ($method -eq "GET" -and ($urlPath -eq "/" -or $urlPath -eq "/ui.html")) {
            $uiPath = Join-Path $PSScriptRoot "ui.html"
            Send-FileResponse -Response $response -FilePath $uiPath -ContentType "text/html; charset=utf-8"
            continue
        }

        if ($method -eq "GET" -and $urlPath -eq "/editor.css") {
            $cssPath = Join-Path $PSScriptRoot "editor.css"
            Send-FileResponse -Response $response -FilePath $cssPath -ContentType "text/css; charset=utf-8"
            continue
        }

        if ($method -eq "GET" -and $urlPath -eq "/editor.js") {
            $jsPath = Join-Path $PSScriptRoot "editor.js"
            Send-FileResponse -Response $response -FilePath $jsPath -ContentType "application/javascript; charset=utf-8"
            continue
        }

        # API: Discover Documents (.odt and .docx)
        if ($method -eq "GET" -and $urlPath -eq "/api/documents") {
            $foundFiles = @()
            $searchRoots = @()

            # Search in working directory, parent, and scratch folder
            $searchRoots += (Get-Location).Path
            $searchRoots += (Split-Path -Parent (Get-Location).Path)
            $searchRoots += "$HOME\.gemini\antigravity\scratch"

            foreach ($root in $searchRoots) {
                if ($root -and (Test-Path $root)) {
                    try {
                        $files = Get-ChildItem -Path $root -Include *.odt, *.docx -File -Recurse -Depth 2 -ErrorAction SilentlyContinue
                        foreach ($f in $files) {
                            $ext = $f.Extension.TrimStart(".").ToLower()
                            $foundFiles += [PSCustomObject]@{
                                name = $f.Name
                                path = $f.FullName
                                format = $ext
                                size = $f.Length
                                lastModified = $f.LastWriteTime.ToString("o")
                            }
                        }
                    } catch {}
                }
            }

            foreach ($shallow in @("$HOME\Documents", "$HOME\Desktop")) {
                if ($shallow -and (Test-Path $shallow)) {
                    try {
                        $files = Get-ChildItem -Path $shallow -Include *.odt, *.docx -File -ErrorAction SilentlyContinue
                        foreach ($f in $files) {
                            $ext = $f.Extension.TrimStart(".").ToLower()
                            $foundFiles += [PSCustomObject]@{
                                name = $f.Name
                                path = $f.FullName
                                format = $ext
                                size = $f.Length
                                lastModified = $f.LastWriteTime.ToString("o")
                            }
                        }
                    } catch {}
                }
            }

            # Deduplicate by path
            $uniqueList = @($foundFiles | Sort-Object path -Unique)
            Send-JsonResponse -Response $response -Data @{ files = $uniqueList }
            continue
        }

        # API: Read Document
        if ($method -eq "GET" -and $urlPath -eq "/api/document") {
            $docPath = $request.QueryString["path"]
            if (-not $docPath) {
                Send-JsonResponse -Response $response -Data @{ error = "Missing 'path' query parameter" } -StatusCode 400
                continue
            }

            if (-not (Test-Path $docPath)) {
                Send-JsonResponse -Response $response -Data @{ error = "File not found: $docPath" } -StatusCode 404
                continue
            }

            try {
                $doc = Convert-DocumentToHtml -DocPath $docPath
                Send-JsonResponse -Response $response -Data @{
                    path = $docPath
                    title = $doc.title
                    format = $doc.format
                    html = $doc.html
                    styles = $doc.styles
                }
            } catch {
                Send-JsonResponse -Response $response -Data @{ error = "Failed to parse document: $($_.Exception.Message)" } -StatusCode 500
            }
            continue
        }

        # API: Save or Create Document
        if ($method -eq "POST" -and $urlPath -eq "/api/document") {
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()

            if (-not $bodyText) {
                Send-JsonResponse -Response $response -Data @{ error = "Empty request body" } -StatusCode 400
                continue
            }

            $body = $bodyText | ConvertFrom-Json
            $docPath = $body.path

            if (-not $docPath) {
                Send-JsonResponse -Response $response -Data @{ error = "Missing 'path' property in request body" } -StatusCode 400
                continue
            }

            $ext = [System.IO.Path]::GetExtension($docPath).ToLower()
            if ($ext -notin @(".odt", ".docx")) {
                Send-JsonResponse -Response $response -Data @{ error = "Unsupported format '$ext'. Expected .odt or .docx" } -StatusCode 400
                continue
            }

            # Create new document request
            if ($body.createNew -eq $true) {
                try {
                    $docTitle = if ($body.title) { $body.title } else { [System.IO.Path]::GetFileNameWithoutExtension($docPath) }
                    New-BlankDocument -DocPath $docPath -Title $docTitle | Out-Null
                    if ($body.html) {
                        Convert-HtmlToDocument -DocPath $docPath -HtmlContent $body.html | Out-Null
                    }
                    Send-JsonResponse -Response $response -Data @{
                        success = $true
                        path = $docPath
                        format = $ext.TrimStart(".")
                        message = "Created new $ext document"
                    }
                } catch {
                    Send-JsonResponse -Response $response -Data @{ error = "Failed to create new document: $($_.Exception.Message)" } -StatusCode 500
                }
                continue
            }

            # Save existing document
            if ($null -eq $body.html) {
                Send-JsonResponse -Response $response -Data @{ error = "Missing 'html' content to save" } -StatusCode 400
                continue
            }

            try {
                Convert-HtmlToDocument -DocPath $docPath -HtmlContent $body.html | Out-Null
                $fileInfo = Get-Item $docPath
                Send-JsonResponse -Response $response -Data @{
                    success = $true
                    path = $docPath
                    format = $ext.TrimStart(".")
                    bytesWritten = $fileInfo.Length
                    lastModified = $fileInfo.LastWriteTime.ToString("o")
                }
            } catch {
                Send-JsonResponse -Response $response -Data @{ error = "Failed to save document: $($_.Exception.Message)" } -StatusCode 500
            }
            continue
        }

        # API: Convert Format / Export
        if ($method -eq "POST" -and $urlPath -eq "/api/convert") {
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()

            if (-not $bodyText) {
                Send-JsonResponse -Response $response -Data @{ error = "Empty request body" } -StatusCode 400
                continue
            }

            $body = $bodyText | ConvertFrom-Json
            $targetPath = $body.targetPath

            if (-not $targetPath) {
                Send-JsonResponse -Response $response -Data @{ error = "Missing 'targetPath' in request body" } -StatusCode 400
                continue
            }

            try {
                Convert-DocumentFormat -SourcePath $body.sourcePath -TargetPath $targetPath -HtmlContent $body.html | Out-Null
                $targetExt = [System.IO.Path]::GetExtension($targetPath).TrimStart(".").ToLower()
                Send-JsonResponse -Response $response -Data @{
                    success = $true
                    targetPath = $targetPath
                    format = $targetExt
                    bytesWritten = (Get-Item $targetPath).Length
                }
            } catch {
                Send-JsonResponse -Response $response -Data @{ error = "Failed to convert document: $($_.Exception.Message)" } -StatusCode 500
            }
            continue
        }

        # Unhandled routes
        Send-JsonResponse -Response $response -Data @{ error = "Not found" } -StatusCode 404

    } catch {
        if ($running) {
            Write-Host "Server request handling error: $($_.Exception.Message)"
        }
    }
}

try {
    $listener.Stop()
    $listener.Close()
} catch {}
Write-Host "Antigravity Document Editor server stopped."
