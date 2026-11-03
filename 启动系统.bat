@echo off
chcp 65001 >nul
title OrderSystem Service Guard
cd /d "%~dp0"
if not exist "data" mkdir "data"
if not exist "images" mkdir "images"
if not exist "inbox" mkdir "inbox"
if not exist "data\backups" mkdir "data\backups"

rem If another valid instance is already running, open it instead of spawning a duplicate.
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $r=Invoke-WebRequest -UseBasicParsing -TimeoutSec 2 -Method Post -ContentType 'application/json' -Body '[]' http://127.0.0.1:8765/api/ping; $expected=[char]34+'pong'+[char]34; if ($r.StatusCode -eq 200 -and $r.Content -eq $expected) { exit 0 }; exit 2 } catch { if ($_.Exception.Response) { exit 2 }; exit 1 }"
set "PROBE=%ERRORLEVEL%"
if "%PROBE%"=="0" (
  start "" "http://127.0.0.1:8765/"
  exit /b 0
)
if "%PROBE%"=="2" (
  echo Port 8765 is already used by another service. Close it or resolve the port conflict first.
  pause
  exit /b 2
)

set "PY="
if exist "%~dp0.venv\Scripts\pythonw.exe" (
  set "PY=%~dp0.venv\Scripts\pythonw.exe"
)
if not defined PY if exist "%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\pythonw.exe" (
  set "PY=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\pythonw.exe"
)
if not defined PY if exist "%LOCALAPPDATA%\.workbuddy\binaries\python\envs\default\Scripts\pythonw.exe" (
  set "PY=%LOCALAPPDATA%\.workbuddy\binaries\python\envs\default\Scripts\pythonw.exe"
)
if not defined PY (
  for /f "delims=" %%P in ('where pythonw 2^>nul') do if not defined PY set "PY=%%P"
)
if not defined PY (
  for /f "delims=" %%P in ('where python 2^>nul') do if not defined PY set "PY=%%P"
)
if not defined PY (
  echo Python was not found. Run the install script or install Python 3.9+ with PATH enabled.
  pause
  exit /b 1
)

:loop
echo %date% %time% Service start >> "data\guard.log"
"%PY%" api_server.py
set "EXITCODE=%ERRORLEVEL%"
if "%EXITCODE%"=="0" exit /b 0
echo %date% %time% Service exited with code %EXITCODE%; restarting in 3 seconds >> "data\guard.log"
timeout /t 3 /nobreak >nul
goto loop
