@echo off
set tm=%time: =0%
set dt=%date:/=%
set BACKUP_LOG="%MYDIRPATH_DESKTOP%\backup_executed_at_%dt:~2,6%-%tm:~0,2%%tm:~3,2%%tm:~6,2%.log"
set CUR_DIR=%~dp0
::echo %BACKUP_LOG%
::echo %CUR_DIR%

if "%~1" == "" (
    set COMMITMSG=Add backup.
) else (
    set COMMITMSG=%~1
)
::echo %COMMITMSG%

cd %CUR_DIR%
echo %date% %time% >> %BACKUP_LOG% 2>>&1
git add -u >> %BACKUP_LOG% 2>>&1
git commit -m "%COMMITMSG%" >> %BACKUP_LOG% 2>>&1
echo. >> %BACKUP_LOG% 2>>&1

