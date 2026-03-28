---
name: powershell-cors-proxy
description: Server HTTP local în PowerShell (System.Net.HttpListener) care servește fișiere statice și proxiază cereri către API-uri externe (ex: Anthropic) ocolind restricțiile CORS din browser.
type: project
---

# PowerShell Local HTTP Server cu CORS Proxy

Server HTTP minimal în PowerShell pur — fără Node.js, fără Python, fără dependențe externe.
Servește fișiere statice și proxiază cereri API pentru a ocoli CORS.

## Cazul de utilizare

Browserul blochează cererile directe către `api.anthropic.com` din `localhost` (CORS policy).
Soluție: PowerShell ascultă pe port 8080, servește `index.html`, și proxiază `/api/ocr` → Anthropic.

## Server complet (server.ps1)

```powershell
$port = 8080
$dir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()

while ($listener.IsListening) {
  try {
    $ctx = $listener.GetContext()
    $req = $ctx.Request
    $res = $ctx.Response

    if ($req.Url.AbsolutePath -eq '/api/ocr' -and $req.HttpMethod -eq 'POST') {
      # Proxy către Anthropic
      # API key din query string ?k=... (evită coruperea în .NET headers)
      $query  = $req.Url.Query.TrimStart('?')
      $params = @{}
      foreach ($pair in $query.Split('&')) {
        $parts = $pair.Split('=', 2)
        if ($parts.Length -eq 2) {
          $params[$parts[0]] = [System.Uri]::UnescapeDataString($parts[1])
        }
      }
      $apiKey = $params['k']
      $beta   = $params['beta']

      # Citește body ca bytes brute (fără parsare JSON — evită erori pe payload mare)
      $ms = New-Object System.IO.MemoryStream
      $req.InputStream.CopyTo($ms)
      $bodyBytes = $ms.ToArray()
      $ms.Close()

      $wreq = [System.Net.HttpWebRequest]::Create('https://api.anthropic.com/v1/messages')
      $wreq.Method       = 'POST'
      $wreq.ContentType  = 'application/json; charset=utf-8'
      $wreq.ContentLength = $bodyBytes.Length
      $wreq.Headers.Add('x-api-key', $apiKey)
      $wreq.Headers.Add('anthropic-version', '2023-06-01')
      if ($beta) { $wreq.Headers.Add('anthropic-beta', $beta) }

      $ws = $wreq.GetRequestStream()
      $ws.Write($bodyBytes, 0, $bodyBytes.Length)
      $ws.Close()

      try {
        $wresp    = $wreq.GetResponse()
        $rms      = New-Object System.IO.MemoryStream
        $wresp.GetResponseStream().CopyTo($rms)
        $wresp.Close()
        $respBytes = $rms.ToArray()
        $res.StatusCode = 200
      } catch [System.Net.WebException] {
        $errRsp   = $_.Exception.Response
        $rms      = New-Object System.IO.MemoryStream
        $errRsp.GetResponseStream().CopyTo($rms)
        $errRsp.Close()
        $respBytes = $rms.ToArray()
        $res.StatusCode = [int]$errRsp.StatusCode
      }

      $res.ContentType     = 'application/json; charset=utf-8'
      $res.ContentLength64 = $respBytes.Length
      $res.OutputStream.Write($respBytes, 0, $respBytes.Length)

    } else {
      # Servește index.html pentru orice altă rută
      $file    = Join-Path $dir "index.html"
      $content = [System.IO.File]::ReadAllBytes($file)
      $res.ContentType     = "text/html; charset=utf-8"
      $res.ContentLength64 = $content.Length
      $res.OutputStream.Write($content, 0, $content.Length)
    }

    $res.OutputStream.Close()
  } catch { }
}
```

## Launcher (run.bat)

```batch
@echo off
cd /d "%~dp0"
timeout /t 1 /nobreak >nul
start "" "http://localhost:8080"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
```

## Lecții învățate

1. **API key în query string, nu în header** — .NET HttpListener corupea valoarea headerului `x-api-key` la caractere speciale
2. **Body ca bytes brute** — parsarea JSON a body-ului în PowerShell eșua pe payload-uri mari cu Base64
3. **`-ExecutionPolicy Bypass`** — necesar pe mașini cu politici restrictive
4. **Erori HTTP** — `WebException` trebuie prinsă separat pentru a returna statusul corect (4xx/5xx)
