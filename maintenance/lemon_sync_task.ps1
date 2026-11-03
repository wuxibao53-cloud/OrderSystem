# LemonCloud stock sync task: POST /api/lemon_sync after service is up. ASCII only.
# curl.exe is used instead of Invoke-RestMethod because the system proxy on this
# machine intercepts 127.0.0.1 requests; curl does not use the system proxy.
$log = 'D:\QMDownload\OrderSystem\data\lemon\task.log'
New-Item -ItemType Directory -Force -Path (Split-Path $log) | Out-Null
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$bodyFile = Join-Path $env:TEMP 'lemon_sync_body.json'
Set-Content -Path $bodyFile -Value '["auto"]' -Encoding Ascii
$ok = $false
for ($i = 0; $i -lt 18; $i++) {
  $out = & curl.exe -s -m 300 -X POST -H 'Content-Type: application/json' `
    -d "@$bodyFile" 'http://127.0.0.1:8765/api/lemon_sync_run' 2>$null
  if ($LASTEXITCODE -eq 0 -and $out) {
    try {
      $r = $out | ConvertFrom-Json
      if ($r.ok) {
        Add-Content -Path $log -Encoding UTF8 -Value (
          "{0} OK checked={1} diffs={2} fixes={3} report={4}" -f `
            $stamp, $r.checked, $r.diffs, @($r.fixes).Count, $r.report)
        $ok = $true
        break
      }
      Add-Content -Path $log -Encoding UTF8 -Value (
        "{0} ERROR from api: {1}" -f $stamp, $r.msg)
      $ok = $true
      break
    } catch {
      Add-Content -Path $log -Encoding UTF8 -Value (
        "{0} ERROR bad response: {1}" -f $stamp, ($out -join ' ').Substring(0, 200))
      $ok = $true
      break
    }
  }
  Start-Sleep -Seconds 10
}
if (-not $ok) {
  Add-Content -Path $log -Encoding UTF8 -Value "$stamp FAIL service not reachable after 18 retries"
}
