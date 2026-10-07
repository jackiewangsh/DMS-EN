@echo off
chcp 936 >nul
cd /d "%~dp0"
echo ========================================
echo   DMS Dev Server
echo ========================================
echo.

REM === Step 1: Create venv if missing ===
if not exist "venv\Scripts\python.exe" goto novenv
goto checkpip
:novenv
echo [DMS] Creating virtual environment...
python -m venv venv
if errorlevel 1 goto venvfail
echo [DMS] venv created OK
goto checkpip
:venvfail
echo [DMS] venv creation failed. Install Python 3.8+ first.
pause
exit /b 1

REM === Step 2: Check ALL required dependencies ===
:checkpip
venv\Scripts\python -c "import flask,pymysql,jieba,pdfplumber,docx,openpyxl,waitress,olefile,reportlab,apscheduler,dotenv,pptx,xlrd,dbutils,win32api" 2>nul
if errorlevel 1 goto pipinstall
echo [DMS] All dependencies OK, skipping pip install.
goto dbsection

REM === Step 3: Install dependencies (local wheels first, then online as fallback) ===
:pipinstall
REM === First option: Install from local wheels ===
echo [DMS] Installing from local wheels...
venv\Scripts\pip install -q -r requirements.txt --no-index --find-links=system\wheels
if not errorlevel 1 goto wheelsok
echo [DMS] Local wheels incomplete, trying PyPI...

REM === Second option: PyPI online ===
venv\Scripts\pip install -q -r requirements.txt
if not errorlevel 1 goto wheelsok

echo [DMS] PyPI failed, trying Tsinghua mirror...
venv\Scripts\pip install -q -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
if not errorlevel 1 goto wheelsok

REM === Third option: Individual wheels as last fallback ===
echo [DMS] Online failed, trying individual wheels...
set WHEELDIR=system\wheels
for %%f in (%WHEELDIR%\*.whl) do (
    echo   Installing %%~nxf...
    venv\Scripts\pip install -q "%%f" --no-deps 2>nul
)
for %%f in (%WHEELDIR%\*.whl) do (
    echo   Resolving deps for %%~nxf...
    venv\Scripts\pip install -q "%%f" 2>&1
)
echo [DMS] Wheel install complete

:wheelsok
echo [DMS] Dependencies installed OK

REM === Step 4: Verify dependencies and show failures ===
echo [DMS] Verifying dependencies...
set FAILDEPS=
for %%m in (flask pymysql jieba pdfplumber docx openpyxl waitress olefile reportlab apscheduler dotenv pptx xlrd dbutils win32api) do (
    venv\Scripts\python -c "import %%m" 2>nul
    if errorlevel 1 set FAILDEPS=!FAILDEPS! %%m
)
if defined FAILDEPS (
    echo [DMS] WARNING: Missing dependencies:!FAILDEPS!
    echo [DMS] Please install manually: venv\Scripts\pip install !FAILDEPS!
    echo.
) else (
    echo [DMS] All dependencies verified OK
)

REM === Step 5: Database setup ===
:dbsection
echo.
echo [DMS] Database setup...
echo   Enter MariaDB root password (press Enter if none):
set /p DBPASS=   Password: 

echo [DMS] Testing database connection...
venv\Scripts\python -c "import pymysql; c=pymysql.connect(host='localhost',port=3306,user='root',password='%DBPASS%'); c.close()" 2>nul
if errorlevel 1 (
    echo [DMS] ERROR: Cannot connect to MariaDB. Check if running.
    pause
    exit /b 1
)

echo [DMS] Creating database if not exists...
echo CREATE DATABASE IF NOT EXISTS dms DEFAULT CHARACTER SET utf8mb4; | mysql -u root --password=%DBPASS% 2>nul

echo [DMS] Importing backup data...
if exist "data\dms_backup.sql" (
    mysql -u root --password=%DBPASS% dms < data\dms_backup.sql 2>nul
    if errorlevel 1 (
        echo [DMS] Backup import failed or already exists.
    ) else (
        echo [DMS] Backup imported successfully.
    )
) else (
    echo [DMS] No backup file (data\dms_backup.sql), skipping import.
    echo [DMS] Run python _backup.py to create backup first.
)
echo [DMS] Database setup complete.
echo.

REM === Dev mode: disable scheduled tasks to avoid watchdog interference ===
echo [DMS] Disabling scheduled tasks (dev mode)...
schtasks /Change /TN "DMS_Watchdog" /DISABLE 2>nul
schtasks /Change /TN "DMS_AutoStart" /DISABLE 2>nul
echo.
echo [DMS] Starting dev server...
echo   URL: http://localhost:5000
echo   Admin: admin / admin123
echo   Press Ctrl+C to stop
echo.
venv\Scripts\python _run.py
pause
exit /b 0