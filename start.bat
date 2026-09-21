@echo off
title Business Card Scanner
cd /d "%~dp0"

:: Kill any old server on port 8080
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8080"') do (
    taskkill /F /PID %%a >nul 2>&1
)

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Python not found. Install Python from https://python.org
    echo     Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b
)

:: Ensure python-docx is installed for Word document generation
python -c "import docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Installing required dependency: python-docx...
    pip install python-docx
)

:: Start server in background without any console window
where pythonw >nul 2>&1
if %errorlevel% equ 0 (
    start "" pythonw.exe "%~dp0server.py" 8080
) else (
    start "" /B python "%~dp0server.py" 8080
)

:: Wait briefly and open the web app in the default browser
start "" /B cmd /c "ping 127.0.0.1 -n 2 >nul && start http://localhost:8080/index.html"

:: Cleanly exit this terminal window so it does not stay open
exit
