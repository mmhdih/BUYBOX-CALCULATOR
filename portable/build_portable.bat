@echo off
rem Builds the portable Windows package (needs Python 3.10+ and internet on THIS machine only).
cd /d "%~dp0.."
python portable\build_portable.py %*
if errorlevel 1 pause & exit /b 1
echo.
echo Done: dist\BuyBoxAnalyzer_portable_win64.zip
pause
