@echo off
REM ============================================================
REM stop_visio_proxy.bat
REM ?? Visio COM ?????
REM ============================================================
echo ???? Visio COM ??...
set LOCK_FILE=D:\codex\projects\draw\scratch\visio_proxy\server.lock
if exist "%LOCK_FILE%" (
    for /f "tokens=1 delims==" %%a in ('type "%LOCK_FILE%"') do set PID=%%a
    if defined PID (
        echo ?????? PID: %PID%
        taskkill /PID %PID% /F 2>nul
    )
)
del "%LOCK_FILE%" 2>nul
echo ?????
pause
