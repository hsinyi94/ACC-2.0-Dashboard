@echo off
REM Windows Task Scheduler / Kiro Hook 執行入口
REM 執行 ACC 2.0 Dashboard 每週更新: build → copy → git commit + push
setlocal
cd /d "%~dp0"

set PY=C:\Users\hsinyih\AppData\Local\Python\bin\python.exe
if not exist "%PY%" set PY=py

"%PY%" scripts\weekly_update.py > "logs\weekly_update_%date:~-4%%date:~3,2%%date:~0,2%.log" 2>&1
set RC=%ERRORLEVEL%

exit /b %RC%
