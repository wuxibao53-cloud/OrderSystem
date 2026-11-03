@echo off
chcp 65001 >nul
cd /d "%~dp0.."
set "PY="
if exist "%~dp0..\.venv\Scripts\pythonw.exe" set "PY=%~dp0..\.venv\Scripts\pythonw.exe"
if not defined PY if exist "%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\pythonw.exe" set "PY=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\pythonw.exe"
if not defined PY for /f "delims=" %%P in ('where pythonw.exe 2^>nul') do if not defined PY set "PY=%%P"
if not defined PY for /f "delims=" %%P in ('where python.exe 2^>nul') do if not defined PY set "PY=%%P"
if not defined PY exit /b 1
"%PY%" "%~dp0daily_backup.py"
