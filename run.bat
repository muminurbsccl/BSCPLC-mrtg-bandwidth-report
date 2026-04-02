@echo off
setlocal enabledelayedexpansion

:: ============================================================
:: MRTG Bandwidth Report - Windows Launcher
:: Auto-detects Tesseract and Poppler from common install paths.
:: No hardcoded user-specific paths.
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
:: Already on PATH?
if not defined TESSERACT where tesseract >nul 2>&1 && goto :tesseract_ok
if not defined TESSERACT (
    echo ERROR: Tesseract OCR not found.
    echo Install from: https://github.com/UB-Mannheim/tesseract/wiki
    pause
    exit /b 1
)
:tesseract_ok

:: --- Poppler: search common locations ---
set "POPPLER="

:: WinGet - any version of oschwartz10612.Poppler
for /d %%D in ("%LocalAppData%\Microsoft\WinGet\Packages\oschwartz10612.Poppler*") do (
    for /d %%V in ("%%D\poppler-*") do (
        if exist "%%V\Library\bin\pdftoppm.exe" set "POPPLER=%%V\Library\bin"
    )
)

:: Chocolatey
if not defined POPPLER if exist "C:\ProgramData\chocolatey\lib\poppler\tools\bin\pdftoppm.exe" (
    set "POPPLER=C:\ProgramData\chocolatey\lib\poppler\tools\bin"
)

:: Manual / scoop / other common locations
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

:: Already on PATH?
if not defined POPPLER where pdftoppm >nul 2>&1 && goto :poppler_ok
if not defined POPPLER (
    echo ERROR: Poppler ^(pdftoppm^) not found.
    echo Install options:
    echo   WinGet:  winget install oschwartz10612.Poppler
    echo   Choco:   choco install poppler
    echo   Manual:  https://github.com/oschwartz10612/poppler-windows/releases
    pause
    exit /b 1
)
:poppler_ok

:: --- Add discovered paths to PATH ---
if defined TESSERACT set "PATH=%PATH%;%TESSERACT%"
if defined POPPLER   set "PATH=%PATH%;%POPPLER%"

:: --- Launch ---
py -3.11 "%~dp0mrtg_bandwidth_report.py"
if errorlevel 1 pause
