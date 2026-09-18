@echo off
title Business Card Scanner
color 0A
echo.
echo  ====================================================
echo    Business Card Scanner -- Local Server Launcher
echo  ====================================================
echo.

:: Kill any old server on port 8080
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8080"') do (
    taskkill /F /PID %%a >nul 2>&1
)

:: Change to the app folder portably
cd /d "%~dp0"

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] Python not found. Install Python from https://python.org
    echo      Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b
)

:: Ensure python-docx is installed for Word document generation
python -c "import docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo  [*] Installing required dependency: python-docx...
    pip install python-docx
)

echo  Server starting at http://localhost:8080
echo.
echo  Camera permission will be saved permanently by Chrome.
echo  Keep this window open while using the app.
echo  Press Ctrl+C to stop.
echo.

:: Open Chrome after 2 seconds
start "" /B cmd /c "timeout /t 2 /nobreak >nul && start chrome http://localhost:8080/index.html"

:: Start custom server with direct Gmail sending support
if exist server.py (
    python server.py 8080
) else (
    python -m http.server 8080
)
pause
