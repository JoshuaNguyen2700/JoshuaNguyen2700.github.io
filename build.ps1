# Builds the single shareable file SLA-Dashboard.html from src\app.html,
# embedding the SheetJS library so the page works fully offline.
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$app  = [IO.File]::ReadAllText("$root\src\app.html")
$lib  = [IO.File]::ReadAllText("$root\vendor\xlsx.full.min.js")
$out  = $app.Replace('/*__SHEETJS__*/', $lib)
[IO.File]::WriteAllText("$root\SLA-Dashboard.html", $out, (New-Object Text.UTF8Encoding($false)))
Write-Output ("Built SLA-Dashboard.html ({0:N0} KB)" -f ((Get-Item "$root\SLA-Dashboard.html").Length / 1KB))
