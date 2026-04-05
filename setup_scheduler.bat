@echo off
:: ============================================================
:: Creates a Windows Task Scheduler task to run auto_report.py
:: daily at 12:05 AM (5 minutes after the email arrives)
:: ============================================================

echo Creating scheduled task: MRTG_Auto_Report
echo Runs daily at 12:05 AM

schtasks /create ^
    /tn "MRTG_Auto_Report" ^
    /tr "\"%~dp0auto_run.bat\"" ^
    /sc daily ^
    /st 00:05 ^
    /rl HIGHEST ^
    /f

if %errorlevel% equ 0 (
    echo.
    echo Task created successfully!
    echo To verify: schtasks /query /tn "MRTG_Auto_Report"
    echo To delete:  schtasks /delete /tn "MRTG_Auto_Report" /f
    echo To run now: schtasks /run /tn "MRTG_Auto_Report"
) else (
    echo ERROR: Failed to create task. Try running this script as Administrator.
)
pause
