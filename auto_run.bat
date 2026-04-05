@echo off
setlocal enabledelayedexpansion

:: ============================================================
:: MRTG Auto Report - Automated Pipeline Launcher
:: Runs auto_report.py: Outlook login → PDF download → OCR → email
:: ============================================================

:: --- Tesseract: search common locations ---
set "TESSERACT="
for %%P in (
    "%ProgramFiles%\Tesseract-OCR"
    "%ProgramFiles(x86)%\Tesseract-OCR"
    "%LocalAppData%\Programs\Tesseract-OCR"
    "C:\tools\tesseract"
) do (
    if not defined TESSERACT if exist "%%~P\tesseract.exe" set "TESSERACT=%%~P"
)
if not defined TESSERACT where tesseract >nul 2>&1 && goto :tesseract_ok
if not defined TESSERACT (
    echo ERROR: Tesseract OCR not found.
    echo Install: winget install UB-Mannheim.TesseractOCR
    pause
    exit /b 1
)
:tesseract_ok

:: --- Poppler: search common locations ---
set "POPPLER="
for /d %%D in ("%LocalAppData%\Microsoft\WinGet\Packages\oschwartz10612.Poppler*") do (
    for /d %%V in ("%%D\poppler-*") do (
        if exist "%%V\Library\bin\pdftoppm.exe" set "POPPLER=%%V\Library\bin"
    )
)
if not defined POPPLER if exist "C:\ProgramData\chocolatey\lib\poppler\tools\bin\pdftoppm.exe" (
    set "POPPLER=C:\ProgramData\chocolatey\lib\poppler\tools\bin"
)
if not defined POPPLER (
    for %%P in (
        "%ProgramFiles%\poppler\bin"
        "%ProgramFiles(x86)%\poppler\bin"
        "C:\tools\poppler\bin"
        "%UserProfile%\poppler\bin"
        "C:\poppler\bin"
    ) do (
        if not defined POPPLER if exist "%%~P\pdftoppm.exe" set "POPPLER=%%~P"
    )
)
if not defined POPPLER where pdftoppm >nul 2>&1 && goto :poppler_ok
if not defined POPPLER (
    echo ERROR: Poppler ^(pdftoppm^) not found.
    echo Install: winget install oschwartz10612.Poppler
    pause
    exit /b 1
)
:poppler_ok

:: --- Check .env file ---
if not exist "%~dp0.env" (
    echo ERROR: .env file not found in %~dp0
    echo Create a .env file with:
    echo   OUTLOOK_EMAIL=your_email@example.com
    echo   OUTLOOK_PASSWORD=your_password
    echo   REPORT_RECIPIENT=recipient@example.com
    echo   TEMPLATE_PATH=D:\path\to\template.xlsx
    pause
    exit /b 1
)

:: --- Add discovered paths to PATH ---
if defined TESSERACT set "PATH=%PATH%;%TESSERACT%"
if defined POPPLER   set "PATH=%PATH%;%POPPLER%"

:: --- Launch auto_report.py ---
echo === MRTG Auto Report Pipeline ===
echo Outlook login, PDF download, OCR report, email delivery
echo.
py -3 "%~dp0auto_report.py"
if errorlevel 1 (
    echo.
    echo Pipeline failed. Check the error above.
    pause
) else (
    echo.
    echo Pipeline complete!
    timeout /t 5
)
