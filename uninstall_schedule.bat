@echo off
setlocal EnableExtensions
schtasks /delete /tn "SCSC_GoogleScholar_Scrape" /f
if errorlevel 1 (
    echo Could not delete the task. It may not exist, or you need Administrator.
) else (
    echo Removed Windows task SCSC_GoogleScholar_Scrape
)
pause
