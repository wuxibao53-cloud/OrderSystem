@echo off
chcp 65001 >nul
set "ROOT=%~dp0"
"%ROOT%记忆层\runtime\node-v20.20.2-win-x64\node.exe" "%ROOT%记忆层\memory.mjs" %*
