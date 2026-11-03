# Register the three Windows scheduled tasks used by the handover package.
# Run this once on the target PC from an elevated PowerShell.
$ErrorActionPreference = "Stop"
$APP = Split-Path -Parent $PSScriptRoot
$bat = Join-Path $APP "启动系统.bat"
$health = Join-Path $APP "maintenance\health_check.ps1"
$hidden = Join-Path $APP "maintenance\health_check_hidden.vbs"
$backup = Join-Path $APP "maintenance\daily_backup.bat"
$who = "$env:USERDOMAIN\$env:USERNAME"

function Register-Task($name, $description, $action, $trigger, $settings) {
    if (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $name -Confirm:$false
    }
    Register-ScheduledTask -TaskName $name -Description $description -Action $action `
        -Trigger $trigger -Settings $settings -User $who -RunLevel Limited | Out-Null
    Write-Host "registered $name"
}

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -StartWhenAvailable `
    -MultipleInstances IgnoreNew

$startAction = New-ScheduledTaskAction -Execute "cmd.exe" `
    -Argument ('/c "{0}"' -f $bat) -WorkingDirectory $APP
$startTrigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
Register-Task "OrderSystem_AutoStart" "Start OrderSystem at user logon" `
    $startAction $startTrigger $settings

$healthAction = New-ScheduledTaskAction -Execute "wscript.exe" `
    -Argument ('"{0}"' -f $hidden) -WorkingDirectory $APP
$healthTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes 5) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
Register-Task "OrderSystem_HealthCheck" "Restart OrderSystem if POST /api/ping fails (windowless via wscript)" `
    $healthAction $healthTrigger $settings

$backupAction = New-ScheduledTaskAction -Execute "cmd.exe" `
    -Argument ('/c "{0}"' -f $backup) -WorkingDirectory $APP
$backupTrigger = New-ScheduledTaskTrigger -Daily -At "18:00"
Register-Task "OrderSystem_DailyBackup" "Daily validated business-data backup" `
    $backupAction $backupTrigger $settings

# 季度恢复演练：隔离目录验证最近备份可完整恢复
$python = Join-Path $env:USERPROFILE ".workbuddy\binaries\python\envs\default\Scripts\python.exe"
if (-not (Test-Path $python)) { $python = "python" }
$verifyAction = New-ScheduledTaskAction -Execute $python `
    -Argument '"{0}"' -f (Join-Path $APP "maintenance\restore_verify.py") -WorkingDirectory $APP
$verifyTrigger = New-ScheduledTaskTrigger -Weekly -At "09:00" -DaysOfWeek Monday `
    -WeeksInterval 12
Register-Task "OrderSystem_RestoreVerify" "Quarterly backup restore drill (isolated temp dir)" `
    $verifyAction $verifyTrigger $settings

# 柠檬云库存同步：开机（延迟 2 分钟）+ 每天 12:00 + 17:35
$lemonAction = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument '-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f `
        (Join-Path $APP "maintenance\lemon_sync_task.ps1") -WorkingDirectory $APP
$bootTrigger = New-ScheduledTaskTrigger -AtStartup
$bootTrigger.Delay = "PT2M"
$noonTrigger = New-ScheduledTaskTrigger -Daily -At "12:00"
$eveningTrigger = New-ScheduledTaskTrigger -Daily -At "17:35"
$lemonSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)
Register-Task "OrderSystem_LemonSync" "LemonCloud stock sync: boot + 12:00 + 17:35 daily" `
    $lemonAction @($bootTrigger, $noonTrigger, $eveningTrigger) $lemonSettings

Write-Host "Done. Tasks: OrderSystem_AutoStart / OrderSystem_HealthCheck / OrderSystem_DailyBackup / OrderSystem_RestoreVerify / OrderSystem_LemonSync"
