<#
.SYNOPSIS
    Antigravity ODT Sidecar HTTP Server.
    Provides web UI and document API endpoints for viewing and editing .odt files.
#>

param(
    [int]$Port = 0,
    [string]$ShutdownToken = ""
)

$ErrorActionPreference = "Continue"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConverterScript = Join-Path $ScriptDir "odt_converter.ps1"
if (Test-Path $ConverterScript) {
    . $ConverterScript
}

# Resolve Port: parameter > env:ANTIGRAVITY_SIDECAR_WEB_PORT > default 58912 (with free port detection)
function Find-AvailablePort {
    param([int]$preferred)
    try {
        $tcp = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $preferred)
        $tcp.Start()
        $tcp.Stop()
        return $preferred
    } catch {
        # Preferred port occupied, find random available port
        $tcp = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
        $tcp.Start()
        $freePort = ($tcp.LocalEndpoint).Port
        $tcp.Stop()
        return $freePort
    }
}

if ($Port -le 0) {
    if ($env:ANTIGRAVITY_SIDECAR_WEB_PORT) {
        $Port = [int]$env:ANTIGRAVITY_SIDECAR_WEB_PORT
    } else {
        $Port = Find-AvailablePort -preferred 58912
    }
}

$Listener = New-Object System.Net.HttpListener
$Prefix = "http://127.0.0.1:$Port/"
$Listener.Prefixes.Add($Prefix)

try {
    $Listener.Start()
    Write-Host "[ODT Sidecar] Listening on $Prefix"
} catch {
    Write-Error "[ODT Sidecar] Failed to start listener on ${Prefix} : $_"
    exit 1
}

# Generate an internal session token for safe shutdown if not passed
if (-not $ShutdownToken) {
    $ShutdownToken = [System.Guid]::NewGuid().ToString("N")
}

function Send-JsonResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [object]$Data,
        [int]$StatusCode = 200
    )
    $json = $Data | ConvertTo-Json -Depth 10
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = "application/json; charset=utf-8"
    $Response.AddHeader("Access-Control-Allow-Origin", "*")
    $Response.AddHeader("Access-Control-Allow-Headers", "*")
    $Response.ContentLength64 = $buffer.Length
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $Response.OutputStream.Close()
}

function Send-FileResponse {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [string]$FilePath,
        [string]$ContentType = "text/html; charset=utf-8"
    )
    if (-not (Test-Path $FilePath)) {
        $Response.StatusCode = 404
        $buffer = [System.Text.Encoding]::UTF8.GetBytes("File not found: $FilePath")
        $Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $Response.OutputStream.Close()
        return
    }
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    $Response.StatusCode = 200
    $Response.ContentType = $ContentType
    $Response.AddHeader("Access-Control-Allow-Origin", "*")
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Validate-OdtPath {
    param([string]$path)
    if ([string]::IsNullOrWhiteSpace($path)) { return $false }
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    return ($ext -eq ".odt")
}

while ($Listener.IsListening) {
    try {
        $context = $Listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $rawUrl = $request.RawUrl
        $path = $request.Url.AbsolutePath
        $method = $request.HttpMethod.ToUpper()

        # Handle CORS preflight
        if ($method -eq "OPTIONS") {
            $response.StatusCode = 200
            $response.AddHeader("Access-Control-Allow-Origin", "*")
            $response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            $response.AddHeader("Access-Control-Allow-Headers", "Content-Type, x-codeium-csrf-token, *")
            $response.OutputStream.Close()
            continue
        }

        # Routes
        if ($path -eq "/api/shutdown") {
            $reqToken = $request.QueryString["token"]
            # Allow shutdown from loopback when token matches or process owner
            if ($reqToken -eq $ShutdownToken -or $request.IsLocal) {
                Send-JsonResponse -Response $response -Data @{ message = "Shutting down" }
                $Listener.Stop()
                break
            } else {
                Send-JsonResponse -Response $response -Data @{ error = "Unauthorized shutdown request" } -StatusCode 403
                continue
            }
        }
        elseif ($path -in @("/", "/ui.html", "/index.html")) {
            $uiFile = Join-Path $ScriptDir "ui.html"
            Send-FileResponse -Response $response -FilePath $uiFile -ContentType "text/html; charset=utf-8"
        }
        elseif ($path -eq "/editor.css") {
            $cssFile = Join-Path $ScriptDir "editor.css"
            Send-FileResponse -Response $response -FilePath $cssFile -ContentType "text/css; charset=utf-8"
        }
        elseif ($path -eq "/editor.js") {
            $jsFile = Join-Path $ScriptDir "editor.js"
            Send-FileResponse -Response $response -FilePath $jsFile -ContentType "application/javascript; charset=utf-8"
        }
        elseif ($path -eq "/api/documents" -and $method -eq "GET") {
            $searchDirs = @(
                (Join-Path $HOME ".gemini\antigravity\scratch"),
                (Get-Location).Path
            )
            $foundFiles = @()
            foreach ($dir in $searchDirs | Select-Object -Unique) {
                if (Test-Path $dir) {
                    $odtFiles = Get-ChildItem -Path $dir -Recurse -Filter "*.odt" -ErrorAction SilentlyContinue | Select-Object -First 30
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
            $uniqueFiles = @($foundFiles | Sort-Object path -Unique)
            Send-JsonResponse -Response $response -Data @{ files = $uniqueFiles }
        }
        elseif ($path -eq "/api/document" -and $method -eq "GET") {
            $targetOdt = $request.QueryString["path"]
            if (-not (Validate-OdtPath -path $targetOdt)) {
                Send-JsonResponse -Response $response -Data @{ error = "Invalid file path or not an .odt document" } -StatusCode 400
                continue
            }
            if (-not (Test-Path $targetOdt)) {
                Send-JsonResponse -Response $response -Data @{ error = "File not found: $targetOdt" } -StatusCode 404
                continue
            }
            try {
                $doc = Convert-OdtToHtml -OdtPath $targetOdt
                Send-JsonResponse -Response $response -Data @{
                    success = $true
                    title = $doc.title
                    html = $doc.html
                    path = $targetOdt
                }
            } catch {
                Send-JsonResponse -Response $response -Data @{ error = $_.Exception.Message } -StatusCode 500
            }
        }
        elseif ($path -eq "/api/document" -and $method -eq "POST") {
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $bodyStr = $reader.ReadToEnd()
            $reader.Close()

            $bodyJson = $null
            try {
                $bodyJson = $bodyStr | ConvertFrom-Json
            } catch {
                Send-JsonResponse -Response $response -Data @{ error = "Invalid JSON body" } -StatusCode 400
                continue
            }

            $targetOdt = if ($bodyJson.path) { $bodyJson.path } else { $request.QueryString["path"] }
            $html = $bodyJson.html

            if (-not (Validate-OdtPath -path $targetOdt)) {
                Send-JsonResponse -Response $response -Data @{ error = "Invalid file path or not an .odt document" } -StatusCode 400
                continue
            }
            if ($null -eq $html) {
                Send-JsonResponse -Response $response -Data @{ error = "Missing 'html' content in request" } -StatusCode 400
                continue
            }

            try {
                $saved = Convert-HtmlToOdt -OdtPath $targetOdt -HtmlContent $html
                Send-JsonResponse -Response $response -Data @{
                    success = [bool]$saved
                    message = "Document updated successfully"
                    path = $targetOdt
                }
            } catch {
                Send-JsonResponse -Response $response -Data @{ error = $_.Exception.Message } -StatusCode 500
            }
        }
        else {
            $response.StatusCode = 404
            $buffer = [System.Text.Encoding]::UTF8.GetBytes("Not found: $rawUrl")
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.OutputStream.Close()
        }
    } catch {
        # Catch connection resets or listener closures
    }
}
