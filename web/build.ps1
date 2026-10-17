# Builds the single-file, offline index.html (JS libraries inlined into web\app.html)
# and Outlook_Connector.bat (web\connector.ps1 wrapped in a double-clickable batch file).
# Run:  powershell -ExecutionPolicy Bypass -File web\build.ps1
$ErrorActionPreference = 'Stop'
$web = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $web
$utf8 = New-Object System.Text.UTF8Encoding($false)

# Outlook_Connector.bat: cmd runs the header, PowerShell runs everything after the marker.
$header = @(
  '@echo off'
  'title Balance Confirmation - Outlook Connector'
  'set "BC_DIR=%~dp0"'
  'set "BC_SELF=%~f0"'
  'powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText($env:BC_SELF); iex $s.Substring($s.IndexOf(''#''+''PSSTART''))"'
  'echo.'
  'pause'
  'exit /b'
  '#PSSTART'
) -join "`r`n"
$ps = [IO.File]::ReadAllText((Join-Path $web 'connector.ps1'), $utf8) -replace "`r?`n", "`r`n"
$bat = $header + "`r`n" + $ps
[IO.File]::WriteAllText((Join-Path $root 'Outlook_Connector.bat'), $bat, (New-Object System.Text.ASCIIEncoding))

$html = [IO.File]::ReadAllText((Join-Path $web 'app.html'), $utf8)
$html = $html.Replace('/*CONNECTOR_BAT*/""', ($bat | ConvertTo-Json -Compress))
$libs = @{ xlsx = 'xlsx.full.min.js'; jspdf = 'jspdf.umd.min.js'; jszip = 'jszip.min.js' }
foreach ($k in $libs.Keys) {
  $js = [IO.File]::ReadAllText((Join-Path $web "lib\$($libs[$k])"), $utf8).Replace('</script', '<\/script')
  $html = $html.Replace("<!--LIB:$k-->", "<script>`n$js`n</script>")
}
# icon embedded so the header logo and favicon also show when index.html is opened from disk
$icon = 'data:image/png;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $root 'icon-192.png')))
$html = $html.Replace('src="icon-192.png"', "src=""$icon""").Replace('href="icon-192.png"', "href=""$icon""")
$out = Join-Path $root 'index.html'
[IO.File]::WriteAllText($out, $html, $utf8)
Write-Host "Built $out ($([math]::Round((Get-Item $out).Length / 1KB)) KB) and Outlook_Connector.bat"
