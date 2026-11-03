$ErrorActionPreference = "Stop"
$log = "D:\QMDownload\OrderSystem\data\restore_task_setup.log"
try {
    $py = "$env:USERPROFILE\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
    $action = New-ScheduledTaskAction -Execute $py `
        -Argument '"D:\QMDownload\OrderSystem\maintenance\restore_verify.py"' `
        -WorkingDirectory "D:\QMDownload\OrderSystem"
    $trigger = New-ScheduledTaskTrigger -Weekly -At "09:00" -DaysOfWeek Monday `
        -WeeksInterval 12
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName "OrderSystem_RestoreVerify" `
        -Description "Quarterly backup restore drill (isolated temp dir)" `
        -Action $action -Trigger $trigger -Settings $settings `
        -User "$env:USERDOMAIN\$env:USERNAME" -RunLevel Limited -Force | Out-Null
    "registered OK" | Out-File $log -Encoding utf8
    exit 0
} catch {
    $_ | Out-String | Out-File $log -Encoding utf8
    exit 1
}
