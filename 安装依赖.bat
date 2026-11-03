@echo off
chcp 65001 >nul
title OrderSystem - Install Dependencies
cd /d "%~dp0"

echo ============================================
echo   OrderSystem - first-time install
echo ============================================
echo.

set "BOOT="
if exist "%~dp0.venv\Scripts\python.exe" set "BOOT=%~dp0.venv\Scripts\python.exe"
if not defined BOOT if exist "%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe" set "BOOT=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not defined BOOT for /f "delims=" %%P in ('where python.exe 2^>nul') do if not defined BOOT set "BOOT=%%P"
if not defined BOOT for /f "delims=" %%P in ('where py.exe 2^>nul') do if not defined BOOT set "BOOT=%%P"

if not defined BOOT (
    echo [x] Python not found.
    echo.
    echo Install Python 3.9+ from https://www.python.org/downloads/
    echo During setup, enable "Add Python to PATH", then run this file again.
    pause
    exit /b 1
)

echo [.] Bootstrap Python:
"%BOOT%" --version
if errorlevel 1 (
    echo [x] This Python cannot run. If it is the Microsoft Store stub, install Python from python.org.
    pause
    exit /b 1
)

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo.
    echo [.] Creating local virtual environment .venv ...
    "%BOOT%" -m venv "%~dp0.venv"
    if errorlevel 1 (
        echo [x] Failed to create .venv. Check that Python 3.9+ is installed correctly.
        pause
        exit /b 1
    )
)

set "PY=%~dp0.venv\Scripts\python.exe"
echo.
echo [.] Installing requirements into .venv (1-3 minutes) ...
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install -r "%~dp0requirements.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo.
    echo Mirror failed, retrying with the official index ...
    "%PY%" -m pip install -r "%~dp0requirements.txt"
)
if errorlevel 1 (
    echo [x] Dependency installation failed. Check the network and retry.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Done. Double-click 启动系统.bat to start.
echo ============================================
pause
