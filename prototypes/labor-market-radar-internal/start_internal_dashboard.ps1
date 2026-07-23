$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$relativeUrl = "/prototypes/labor-market-radar-internal/index.html"
$startPort = 8787
$listener = $null
$port = $null

for ($candidate = $startPort; $candidate -lt ($startPort + 20); $candidate++) {
  try {
    $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Parse("127.0.0.1"), $candidate)
    $listener.Start()
    $port = $candidate
    break
  } catch {
    if ($listener) {
      $listener.Stop()
      $listener = $null
    }
  }
}

if (-not $listener) {
  throw "Не удалось запустить локальный сервер на портах $startPort-$($startPort + 19)."
}

$url = "http://127.0.0.1:$port$relativeUrl"
Write-Host "Внутренний демо-дашборд 1С запущен:"
Write-Host $url
Write-Host ""
Write-Host "Чтобы остановить сервер, закройте это окно или нажмите Ctrl+C."
Start-Process $url

function Get-ContentType([string]$Path) {
  switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
    ".html" { "text/html; charset=utf-8" }
    ".css" { "text/css; charset=utf-8" }
    ".js" { "application/javascript; charset=utf-8" }
    ".json" { "application/json; charset=utf-8" }
    ".png" { "image/png" }
    ".jpg" { "image/jpeg" }
    ".jpeg" { "image/jpeg" }
    ".svg" { "image/svg+xml" }
    default { "application/octet-stream" }
  }
}

while ($true) {
  $client = $listener.AcceptTcpClient()
  try {
    $stream = $client.GetStream()
    $reader = [System.IO.StreamReader]::new($stream)
    $requestLine = $reader.ReadLine()

    if (-not $requestLine) {
      $client.Close()
      continue
    }

    while ($reader.Peek() -ge 0) {
      $line = $reader.ReadLine()
      if ($line -eq "") { break }
    }

    $parts = $requestLine.Split(" ")
    $requestPath = if ($parts.Count -ge 2) { $parts[1] } else { "/" }
    if ($requestPath -eq "/") {
      $requestPath = $relativeUrl
    }
    $requestPath = [System.Uri]::UnescapeDataString(($requestPath -split "\?")[0])
    $requestPath = $requestPath.TrimStart("/")
    $localPath = Join-Path $projectRoot $requestPath
    $fullPath = [System.IO.Path]::GetFullPath($localPath)

    if (-not $fullPath.StartsWith($projectRoot.Path, [System.StringComparison]::OrdinalIgnoreCase)) {
      $body = [System.Text.Encoding]::UTF8.GetBytes("403 Forbidden")
      $header = "HTTP/1.1 403 Forbidden`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n"
      $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
      $stream.Write($headerBytes, 0, $headerBytes.Length)
      $stream.Write($body, 0, $body.Length)
      continue
    }

    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
      $body = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
      $header = "HTTP/1.1 404 Not Found`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n"
      $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
      $stream.Write($headerBytes, 0, $headerBytes.Length)
      $stream.Write($body, 0, $body.Length)
      continue
    }

    $bodyBytes = [System.IO.File]::ReadAllBytes($fullPath)
    $contentType = Get-ContentType $fullPath
    $responseHeader = "HTTP/1.1 200 OK`r`nContent-Type: $contentType`r`nContent-Length: $($bodyBytes.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n`r`n"
    $responseHeaderBytes = [System.Text.Encoding]::ASCII.GetBytes($responseHeader)
    $stream.Write($responseHeaderBytes, 0, $responseHeaderBytes.Length)
    $stream.Write($bodyBytes, 0, $bodyBytes.Length)
  } finally {
    $client.Close()
  }
}
