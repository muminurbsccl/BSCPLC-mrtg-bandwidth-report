@echo off
set "TESSERACT=C:\Program Files\Tesseract-OCR"
set "POPPLER=C:\Users\Mumin\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin"
set "PATH=%PATH%;%TESSERACT%;%POPPLER%"

py -3.11 "%~dp0mrtg_bandwidth_report.py"
