# Outlook Connector for the Vendor Balance Confirmation page.
# Runs a tiny web server on http://127.0.0.1:8765 (reachable only from this PC, no admin rights needed)
# and sends mail through the Outlook desktop app. Every API request must carry the access code
# shown in this window.
$ErrorActionPreference = 'Stop'
$Port = 8765
$Dir = $env:BC_DIR
if (-not $Dir) { $Dir = (Get-Location).Path }
$OutDir = Join-Path $Dir 'output'
$LogPath = Join-Path $OutDir 'sent_log.csv'
$AllowedHosts = @("127.0.0.1:$Port", "localhost:$Port")

$alpha = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'
$bytes = New-Object byte[] 8
(New-Object System.Security.Cryptography.RNGCryptoServiceProvider).GetBytes($bytes)
$Token = -join ($bytes | ForEach-Object { $alpha[$_ % 32] })

$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
try { $listener.Start() }
catch { Write-Host "  Port $Port is busy. Is the Outlook Connector already open in another window?" -ForegroundColor Red; exit 1 }

# When index.html sits next to this file, serve it with the access code built in.
$page = $null
$pagePath = Join-Path $Dir 'index.html'
if (Test-Path $pagePath) {
  $html = [IO.File]::ReadAllText($pagePath, [Text.Encoding]::UTF8)
  $page = [Text.Encoding]::UTF8.GetBytes($html.Insert($html.IndexOf('</head>'), "<script>window.BC_TOKEN='$Token';</script>"))
}

$script:ol = $null
function Get-Outlook {
  if (-not $script:ol) { $script:ol = New-Object -ComObject Outlook.Application }
  return $script:ol
}

# ---------------------------------------------------------------- minimal HTTP
function Read-Request($client) {
  $client.ReceiveTimeout = 3000
  $stream = $client.GetStream()
  $buf = New-Object byte[] 65536
  $ms = New-Object IO.MemoryStream
  $headerEnd = -1
  while ($headerEnd -lt 0) {
    $n = $stream.Read($buf, 0, $buf.Length)
    if ($n -le 0) { return $null }
    $ms.Write($buf, 0, $n)
    $headerEnd = [Text.Encoding]::ASCII.GetString($ms.ToArray()).IndexOf("`r`n`r`n")
    if ($headerEnd -lt 0 -and $ms.Length -gt 65536) { return $null }
  }
  $all = $ms.ToArray()
  $lines = [Text.Encoding]::ASCII.GetString($all, 0, $headerEnd) -split "`r`n"
  $first = $lines[0] -split ' '
  $h = @{}
  foreach ($l in ($lines | Select-Object -Skip 1)) {
    $i = $l.IndexOf(':')
    if ($i -gt 0) { $h[$l.Substring(0, $i).Trim().ToLower()] = $l.Substring($i + 1).Trim() }
  }
  $len = 0
  if ($h['content-length']) { $len = [int]$h['content-length'] }
  $body = New-Object IO.MemoryStream
  $body.Write($all, $headerEnd + 4, $all.Length - $headerEnd - 4)
  $client.ReceiveTimeout = 15000
  while ($body.Length -lt $len) {
    $n = $stream.Read($buf, 0, $buf.Length)
    if ($n -le 0) { break }
    $body.Write($buf, 0, $n)
  }
  return @{ method = $first[0]; path = ($first[1] -split '\?')[0]; headers = $h; stream = $stream
            body = [Text.Encoding]::UTF8.GetString($body.ToArray()) }
}
function Send-Bytes($req, [int]$code, [byte[]]$body, $type, [bool]$cors) {
  $reason = @{ 200 = 'OK'; 204 = 'No Content'; 401 = 'Unauthorized'; 403 = 'Forbidden'; 404 = 'Not Found'; 500 = 'Internal Server Error' }[$code]
  $head = "HTTP/1.1 $code $reason`r`nContent-Type: $type`r`nContent-Length: $($body.Length)`r`nConnection: close`r`nCache-Control: no-store`r`n"
  if ($cors) {
    $head += "Access-Control-Allow-Origin: *`r`nAccess-Control-Allow-Headers: Content-Type, X-Token`r`n"
    $head += "Access-Control-Allow-Methods: GET, POST, OPTIONS`r`nAccess-Control-Allow-Private-Network: true`r`n"
  }
  $hb = [Text.Encoding]::ASCII.GetBytes($head + "`r`n")
  $req.stream.Write($hb, 0, $hb.Length)
  if ($body.Length) { $req.stream.Write($body, 0, $body.Length) }
  $req.stream.Flush()
}
function Send-Json($req, $obj, $code = 200) {
  Send-Bytes $req $code ([Text.Encoding]::UTF8.GetBytes(($obj | ConvertTo-Json -Compress -Depth 5))) 'application/json; charset=utf-8' $true
}

# ---------------------------------------------------------------- Outlook
function Send-Mail($d) {
  $ol = Get-Outlook
  $name = ([string]$d.filename) -replace '[\\/:*?"<>|]', ''
  if (-not $name) { $name = 'letter' }
  if (-not $name.ToLower().EndsWith('.pdf')) { $name += '.pdf' }
  $folder = Join-Path $OutDir ((([string]$d.as_at) -replace '[^0-9-]', ''))
  New-Item -ItemType Directory -Force $folder | Out-Null
  $pdfPath = Join-Path $folder $name
  [IO.File]::WriteAllBytes($pdfPath, [Convert]::FromBase64String($d.pdf))

  $acc = $null
  if ($d.from) {
    foreach ($a in $ol.Session.Accounts) { if ($a.SmtpAddress -ieq $d.from -or $a.DisplayName -ieq $d.from) { $acc = $a } }
    if (-not $acc) { throw "Outlook account '$($d.from)' was not found." }
  }
  $m = $ol.CreateItem(0)
  $m.To = $d.to
  if ($d.cc) { $m.CC = $d.cc }
  $m.Subject = $d.subject
  $m.HTMLBody = $d.body
  [void]$m.Attachments.Add($pdfPath)
  if ($acc) { [void][System.__ComObject].InvokeMember('SendUsingAccount', [System.Reflection.BindingFlags]::SetProperty, $null, $m, @($acc)) }
  switch ($d.action) {
    'Send'  { $m.Send() }
    'Draft' { $m.Save() }
    default { $m.Display($false) }
  }
  return $pdfPath
}
function Write-Log($d, $status, $err, $pdf) {
  New-Item -ItemType Directory -Force $OutDir | Out-Null
  [pscustomobject]@{
    timestamp = (Get-Date -Format 's'); as_at = $d.as_at; vendor_code = $d.vendor_code; vendor_name = $d.vendor_name
    net = $d.net; to = $d.to; cc = $d.cc; mode = $d.mode; action = $d.action; status = $status; error = $err; pdf = $pdf
  } | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Append
}

function Handle($req) {
  # Reject other host names (protects against DNS-rebinding web pages).
  if ($AllowedHosts -notcontains $req.headers['host']) { Send-Json $req @{ ok = $false; error = 'Forbidden host.' } 403; return }
  if ($req.method -eq 'GET' -and ($req.path -eq '/' -or $req.path -eq '/index.html')) {
    if ($page) { Send-Bytes $req 200 $page 'text/html; charset=utf-8' $false }
    else { Send-Bytes $req 404 ([Text.Encoding]::UTF8.GetBytes('index.html not found next to the connector.')) 'text/plain' $false }
    return
  }
  if ($req.method -eq 'OPTIONS') { Send-Bytes $req 204 (New-Object byte[] 0) 'text/plain' $true; return }
  $authorized = ($req.headers['x-token'] -ceq $Token)
  if ($req.path -eq '/ping') { Send-Json $req @{ ok = $true; app = 'bc-outlook-connector'; version = 1; authorized = $authorized }; return }
  if (-not $authorized) { Send-Json $req @{ ok = $false; error = 'Wrong access code.' } 401; return }

  switch ($req.path) {
    '/accounts' {
      # SmtpAddress can be blank for some Exchange accounts; the display name is usually the address.
      $list = @(foreach ($a in (Get-Outlook).Session.Accounts) { if ($a.SmtpAddress) { [string]$a.SmtpAddress } elseif ($a.DisplayName) { [string]$a.DisplayName } })
      Send-Json $req @{ ok = $true; accounts = $list }
    }
    '/send' {
      $d = $req.body | ConvertFrom-Json
      $label = @{ Send = 'Sent'; Draft = 'Draft saved' }[[string]$d.action]
      if (-not $label) { $label = 'Opened' }
      try {
        $pdf = Send-Mail $d
        Write-Log $d $label '' $pdf
        Write-Host ("  {0}  {1} - {2} -> {3}" -f $label, $d.vendor_code, $d.vendor_name, $d.to) -ForegroundColor Green
        Send-Json $req @{ ok = $true; status = $label; pdf = $pdf }
      } catch {
        $msg = $_.Exception.Message
        Write-Log $d 'Failed' $msg ''
        Write-Host ("  Failed  {0} - {1}: {2}" -f $d.vendor_code, $d.vendor_name, $msg) -ForegroundColor Red
        Send-Json $req @{ ok = $false; error = $msg }
      }
    }
    default { Send-Json $req @{ ok = $false; error = 'Unknown request.' } 404 }
  }
}

Write-Host ''
Write-Host '  Vendor Balance Confirmation - Outlook Connector' -ForegroundColor Cyan
Write-Host ''
Write-Host "  Access code :  $Token" -ForegroundColor Yellow
Write-Host "  Address     :  http://127.0.0.1:$Port/"
Write-Host "  Letters/log :  $OutDir"
Write-Host ''
Write-Host '  Keep this window open while sending. Close it when you are done.'
Write-Host ''
if ($page) { if (-not $env:BC_NO_BROWSER) { Start-Process "http://127.0.0.1:$Port/" } }
else { Write-Host '  index.html is not in this folder - open it yourself and enter the access code.' -ForegroundColor DarkYellow }

while ($true) {
  $client = $listener.AcceptTcpClient()
  $req = $null
  try {
    $req = Read-Request $client
    if ($req) { Handle $req }
  } catch {
    if ($req) { try { Send-Json $req @{ ok = $false; error = $_.Exception.Message } 500 } catch {} }
  } finally {
    $client.Close()
  }
}
