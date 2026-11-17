@echo off
setlocal EnableExtensions
cd /d "%~dp0"

REM Set to 0 if you only want a local csv and do not want git push.
set "PUSH_TO_GITHUB=1"

set "LOGDIR=%~dp0logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"
set "LOG=%LOGDIR%\scraper_run.log"

set "SCRAPE_ERR=0"

echo.
echo SCSC Google Scholar scrape
echo Log: %LOG%
echo This usually takes 15-20 minutes.
echo.

echo.>> "%LOG%"
echo ===== %date% %time% =====>> "%LOG%"

set "PY=python"
if exist "%~dp0venv\Scripts\python.exe" set "PY=%~dp0venv\Scripts\python.exe"

"%PY%" --version >> "%LOG%" 2>&1
if errorlevel 1 (
    echo Python not found. Install Python 3 and pip install requests beautifulsoup4 openpyxl
    echo Python not found.>> "%LOG%"
    set "SCRAPE_ERR=1"
    goto :end
)

if exist "New Google Scholar List.csv" (
    copy /Y "New Google Scholar List.csv" "New Google Scholar List.csv.bak" >nul
)

echo Running scraper...
"%PY%" -u scholar_scraper.py %* >> "%LOG%" 2>&1
set "SCRAPE_ERR=%ERRORLEVEL%"
echo Scraper exit code: %SCRAPE_ERR%>> "%LOG%"

"%PY%" -c "import csv,sys; f=open(r'New Google Scholar List.csv',encoding='utf-8-sig'); rows=list(csv.DictReader(f)); n=sum(1 for r in rows if str(r.get('Citations (All)') or '').strip()); print('rows_with_citations', n); sys.exit(0 if n>=50 else 1)" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo Scrape did not produce citation numbers. Restoring the previous csv so GitHub is not wiped.
    echo Restoring previous csv.>> "%LOG%"
    if exist "New Google Scholar List.csv.bak" copy /Y "New Google Scholar List.csv.bak" "New Google Scholar List.csv" >nul
    set "SCRAPE_ERR=1"
    goto :end
)

if not "%PUSH_TO_GITHUB%"=="1" goto :end
if not exist ".git" goto :end

echo Pushing csv to GitHub...
git add "New Google Scholar List.csv" >> "%LOG%" 2>&1
git diff --cached --quiet >> "%LOG%" 2>&1
if errorlevel 1 (
    git commit -m "chore: automated scholar scrape" >> "%LOG%" 2>&1
    git push >> "%LOG%" 2>&1
    if errorlevel 1 (
        echo Git push failed. Csv is updated locally. Check git credentials.
        echo Git push failed.>> "%LOG%"
    ) else (
        echo Pushed New Google Scholar List.csv
    )
) else (
    echo No csv changes to push.
)

:end
echo.
if /I not "%~1"=="/silent" pause
exit /b %SCRAPE_ERR%
