# 维护脚本

- `health_check.ps1`：每 5 分钟 POST `http://127.0.0.1:8765/api/ping`，失败时调用 `启动系统.bat` 拉起服务；日志在 `data\health_check.log`。
- `health_check_hidden.vbs`：计划任务经 `wscript.exe` 运行本文件来隐藏启动 `health_check.ps1`，避免每次检查弹出 PowerShell 黑窗口（`powershell.exe` 是控制台程序，直接挂任务会一闪一闪弹窗）。
- `daily_backup.py` / `daily_backup.bat`：调用 `backup_manager.create_backup()`，把四个业务 JSON 备份到 `data\backups`；日志在 `data\backup_task.log`。
- `install_tasks.ps1`：注册三个任务计划：
  - `OrderSystem_AutoStart`：当前用户登录时启动。
  - `OrderSystem_HealthCheck`：每 5 分钟健康检查。
  - `OrderSystem_DailyBackup`：每天 18:00 静默备份。

在目标电脑上以当前用户运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\maintenance\install_tasks.ps1"
```

若系统提示权限不足，再用管理员 PowerShell 运行同一脚本。任务以当前用户身份运行，不需要管理员权限即可启动本系统和写 `data` 目录。

卸载任务：

```powershell
'OrderSystem_AutoStart','OrderSystem_HealthCheck','OrderSystem_DailyBackup' |
  ForEach-Object { Unregister-ScheduledTask -TaskName $_ -Confirm:$false }
```
