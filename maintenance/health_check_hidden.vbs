' Hidden launcher for maintenance\health_check.ps1 (scheduled task: OrderSystem_HealthCheck)
' The scheduled task runs wscript.exe with this file instead of powershell.exe directly,
' because powershell.exe is a console program and flashes a black window every run.
' wscript.exe is a GUI program, so nothing appears on screen.
Dim fso, sh, appDir, ps1
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
appDir = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
ps1 = fso.BuildPath(appDir, "maintenance\health_check.ps1")
If fso.FileExists(ps1) Then
    sh.CurrentDirectory = appDir
    sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """", 0, False
End If
