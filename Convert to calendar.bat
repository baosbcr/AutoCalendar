@echo off
REM Double-click this to pick a spreadsheet and get an .ics next to it.
cd /d "%~dp0"
python -m autocalendar %*
echo.
pause
