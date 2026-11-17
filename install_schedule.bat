@echo off
setlocal EnableExtensions
cd /d "%~dp0"

REM Creates a Task Scheduler job for Mon/Thu 08:00.
REM Runs only while this user is logged on (locked screen is fine; sign-out is not).
REM No admin / domain password needed.

set "TASK_NAME=SCSC_GoogleScholar_Scrape"
set "BAT=%~dp0run_scraper_silent.bat"

schtasks /create /tn "%TASK_NAME%" /tr "\"%BAT%\"" /sc weekly /d MON,THU /st 08:00 /rl LIMITED /f
if errorlevel 1 (
    echo Failed to create the task. Try running this bat as Administrator.
    pause
    exit /b 1
)

echo.
echo Created Windows task "%TASK_NAME%"
echo Runs Monday and Thursday at 08:00 while you are logged on.
echo Lock the screen is OK. Sign out / reboot without logging in is not.
echo.
echo Remove it: uninstall_schedule.bat
echo.
pause
