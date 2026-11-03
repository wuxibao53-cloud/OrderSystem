# 柠檬云库存同步任务：开机（延迟 2 分钟）+ 每天 12:00 + 17:35
$ErrorActionPreference = "Stop"
$log = "D:\QMDownload\OrderSystem\data\lemon\register.log"
try {
    New-Item -ItemType Directory -Force -Path (Split-Path $log) | Out-Null
    $action = New-ScheduledTaskAction -Execute "powershell.exe" `
        -Argument '-NoProfile -ExecutionPolicy Bypass -File "D:\QMDownload\OrderSystem\maintenance\lemon_sync_task.ps1"' `
        -WorkingDirectory "D:\QMDownload\OrderSystem"
    $boot = New-ScheduledTaskTrigger -AtStartup
    $boot.Delay = "PT2M"
    $noon = New-ScheduledTaskTrigger -Daily -At "12:00"
    $evening = New-ScheduledTaskTrigger -Daily -At "17:35"
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1)
    Register-ScheduledTask -TaskName "OrderSystem_LemonSync" `
        -Description "LemonCloud stock sync: boot + 12:00 + 17:35 daily" `
        -Action $action -Trigger $boot, $noon, $evening -Settings $settings `
        -User "$env:USERDOMAIN\$env:USERNAME" -RunLevel Limited -Force | Out-Null
    "registered OK $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File $log -Encoding utf8
    exit 0
} catch {
    $_ | Out-String | Out-File $log -Encoding utf8
    exit 1
}
