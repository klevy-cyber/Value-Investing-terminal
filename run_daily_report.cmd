@echo off
rem Daily special-situations workbook (Event Monitor layout), run by Windows Task Scheduler.
rem Writes reports\daily\special_situations_monitor_YYYY-MM-DD.xlsx; the log goes to reports\daily\daily_log.txt.
cd /d "%~dp0"
if not exist reports\daily mkdir reports\daily
set "PYTHONIOENCODING=utf-8"
set "PYTHON=C:\Users\KLevy\AppData\Local\Python\pythoncore-3.15-64\python.exe"
if not exist "%PYTHON%" set "PYTHON=py"
echo ==== %date% %time% ====>> reports\daily\daily_log.txt
"%PYTHON%" daily_report.py %* >> reports\daily\daily_log.txt 2>&1
