@echo off
set DST_DIR_UPATH=~/_scp_from_win
set USER=user
set HOST=XXX.XXX.XXX.XXX
set PASSWORD=password

if "%~1"=="" (
    echo �������w�肵�Ă��������B�����𒆒f���܂��B
    pause
    exit /b
)
set SRCPATH_WPATH=%~1
echo %SRCPATH_WPATH%
for /f "usebackq" %%A in (`wsl wslpath -a "%SRCPATH_WPATH%"`) do set SRCPATH_UPATH=%%A
echo %SRCPATH_UPATH%

wsl expect -c "spawn scp -r %SRCPATH_UPATH% %USER%@%HOST%:%DST_DIR_UPATH% ; expect password: ; send %PASSWORD%\r ; expect $ ; interact"
pause
