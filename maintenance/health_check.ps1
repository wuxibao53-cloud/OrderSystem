$ErrorActionPreference = "SilentlyContinue"
$APP = Split-Path -Parent $PSScriptRoot
$log = Join-Path $APP "data\health_check.log"
$dataDir = Join-Path $APP "data"
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }

$failFile = Join-Path $dataDir "health_fail.count"
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
function Write-Log([string]$message) {
    Add-Content -LiteralPath $log -Value ("{0} {1}" -f $stamp, $message) -Encoding UTF8
}

try {
    $response = Invoke-WebRequest -UseBasicParsing -TimeoutSec 4 -Method Post `
        -ContentType "application/json" -Body "[]" "http://127.0.0.1:8765/api/ping"
    if ($response.StatusCode -eq 200 -and $response.Content.Trim() -eq '"pong"') {
        if (Test-Path -LiteralPath $failFile) { Remove-Item -LiteralPath $failFile -Force }
        exit 0
    }
} catch {
}

$fails = 1
if (Test-Path -LiteralPath $failFile) { $fails = [int](Get-Content -LiteralPath $failFile -ErrorAction SilentlyContinue | Select-Object -First 1) + 1 }
Set-Content -LiteralPath $failFile -Value $fails
if ($fails -lt 3) { Write-Log "health check failed ($fails/3); waiting"; exit 0 }
Remove-Item -LiteralPath $failFile -Force -ErrorAction SilentlyContinue
Write-Log "health check failed 3 times; starting the service"
$launcher = Join-Path $APP "启动系统.bat"
if (Test-Path $launcher) {
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c", ('"{0}"' -f $launcher) `
        -WorkingDirectory $APP -WindowStyle Hidden
} else {
    Write-Log "launcher not found: $launcher"
}
