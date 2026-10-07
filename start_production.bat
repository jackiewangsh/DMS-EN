@echo off
setlocal enabledelayedexpansion
chcp 936 >nul
cd /d "%~dp0"
echo ========================================
echo   DMS Production Server
echo ========================================
echo.

REM === Step 1: Create venv if missing ===
REM Check both python.exe AND pyvenv.cfg (incomplete venv causes "No pyvenv.cfg file" error)
if not exist "venv\Scripts\python.exe" goto novenv
if not exist "venv\pyvenv.cfg" goto novenv
REM if not exist "venv\Lib\os.py" goto novenv
goto checkpip
:novenv
echo [DMS] Creating virtual environment...
if exist venv (
    echo [DMS] Removing incomplete venv directory...
    rmdir /s /q venv
)
python -m venv venv
if errorlevel 1 goto venvfail
if not exist "venv\pyvenv.cfg" (
    echo [DMS] venv created but pyvenv.cfg missing! Trying again...
    rmdir /s /q venv
    python -m venv venv
    if errorlevel 1 goto venvfail
)
echo [DMS] venv created OK
goto checkpip
:venvfail
echo [DMS] venv creation failed. Please install Python 3.8+
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

REM === Step 5: Database setup (only on first run) ===
:dbsection
echo.

REM Check if already configured (marker file)
if exist "data\.db_configured" goto startserver

REM Try connecting with empty password first (common for local dev)
set "DBPASS="
set DB_TEST=0
for /f "tokens=*" %%i in ('mysql -u root --password= -e "SHOW DATABASES LIKE 'dms'" 2^>nul') do (
    if "%%i"=="dms" set DB_TEST=1
)
if %DB_TEST%==1 goto dbfound_empty

REM Empty password didn't work - prompt for password (MariaDB often requires one)
echo [DMS] Database not found. Enter MariaDB root password:
set /p DBPASS=   Password (leave blank if none): 
set DB_TEST=0
for /f "tokens=*" %%i in ('mysql -u root --password=%DBPASS% -e "SHOW DATABASES LIKE 'dms'" 2^>nul') do (
    if "%%i"=="dms" set DB_TEST=1
)
if %DB_TEST%==0 (
    echo [DMS] ERROR: Cannot connect to MariaDB. Check password or start MariaDB service.
    pause
    exit /b 1
)
REM Database found with password
goto dbfound_pass

REM === Database found with empty password ===
:dbfound_empty
REM Check if dms database has any tables with data
set DB_HAS_TABLES=0
for /f "skip=1 tokens=*" %%i in ('mysql -u root dms -e "SELECT COUNT(*) FROM files" 2^>nul') do (
    set DB_HAS_TABLES=1
)
if %DB_HAS_TABLES%==1 (
    echo [DMS] Database already configured, skipping setup.
    echo. > "data\.db_configured"
    goto startserver
) else (
    REM Database exists but no data - run first-time setup
    goto dbfirstrun_empty
)

REM === Database found with password ===
:dbfound_pass
REM Check if dms database has any tables with data
set DB_HAS_TABLES=0
for /f "skip=1 tokens=*" %%i in ('mysql -u root --password=%DBPASS% dms -e "SELECT COUNT(*) FROM files" 2^>nul') do (
    set DB_HAS_TABLES=1
)
if %DB_HAS_TABLES%==1 (
    echo [DMS] Database already configured, skipping setup.
    echo. > "data\.db_configured"
    goto startserver
) else (
    REM Database exists but no data - run first-time setup
    goto dbfirstrun
)

REM === First-time: create database and import backup (empty password) ===
:dbfirstrun_empty
echo [DMS] Creating database...
mysql -u root -e "CREATE DATABASE IF NOT EXISTS dms DEFAULT CHARACTER SET utf8mb4" 2>nul
if not errorlevel 1 (
    echo [DMS] Database created.
) else (
    echo [DMS] WARNING: Could not create database.
)
if exist "data\dms_backup.sql" (
    echo [DMS] Importing backup data (first-time setup)...
    mysql -u root dms < data\dms_backup.sql 2>nul
    if not errorlevel 1 (
        echo [DMS] Backup imported successfully.
    ) else (
        echo [DMS] Backup import skipped (tables may already exist).
    )
) else (
    echo [DMS] No backup file found (data\dms_backup.sql).
    echo [DMS] Run python _backup.py to create backup first.
)
echo [DMS] Database setup complete.
echo. > "data\.db_configured"
echo.

REM === 创建计划任务（开机自启 + Watchdog）===
echo [DMS] 配置计划任务...

REM 开机自启任务（SYSTEM 账户，无窗口）
schtasks /Delete /TN "DMS_AutoStart" /F >nul 2>&1
schtasks /Query /TN "DMS_AutoStart" >nul 2>&1
if errorlevel 1 (
    schtasks /Create /TN "DMS_AutoStart" /TR "\"%CD%\venv\Scripts\python.exe\" \"%CD%\_run.py\"" /SC ONSTART /RU SYSTEM /F
    echo [DMS] 已创建开机自启任务（DMS_AutoStart）
) else (
    echo [DMS] 开机自启任务已存在（DMS_AutoStart）
)

REM Watchdog 任务（每 60 秒检查一次，SYSTEM 账户）
schtasks /Delete /TN "DMS_Watchdog" /F >nul 2>&1
schtasks /Query /TN "DMS_Watchdog" >nul 2>&1
if errorlevel 1 (
    schtasks /Create /TN "DMS_Watchdog" /TR "\"%CD%\venv\Scripts\python.exe\" \"%CD%\watchdog.py\"" /SC MINUTE /MO 1 /RU SYSTEM /F
    echo [DMS] 已创建 Watchdog 任务（DMS_Watchdog，每 60 秒）
) else (
    echo [DMS] Watchdog 任务已存在（DMS_Watchdog）
)

echo.
goto startserver

:startserver
echo.
echo [DMS] Starting production server...
echo   URL: http://localhost:5000
echo   Admin: admin / admin123
echo   Watchdog: Enabled (auto-restart on crash)
echo.

:restart
echo [%date% %time%] Starting DMS...
venv\Scripts\python _run.py
set EXITCODE=%errorlevel%
if %EXITCODE% equ 0 goto ended
echo [%date% %time%] DMS crashed (exit code %EXITCODE%), restarting in 5s...
timeout /t 5 /nobreak >nul
goto restart

:ended
echo [%date% %time%] DMS stopped normally.
pause
