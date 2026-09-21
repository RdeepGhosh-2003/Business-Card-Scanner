@echo off
title Stop Business Card Scanner
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8080"') do (
    taskkill /F /PID %%a >nul 2>&1
)
echo Server on port 8080 stopped successfully.
ping 127.0.0.1 -n 2 >nul
exit
