@echo off
chcp 65001 >nul 2>&1
echo.
echo  Closing DMS Service...
echo.

taskkill /F /IM python.exe >nul 2>&1
if %errorlevel% equ 0 (
    echo  DMS Closed Successfully
) else (
    echo  Not found the python process running
)

echo.
pause
