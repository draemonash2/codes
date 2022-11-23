@echo off
set DST_DIR_UPATH=~/_scp_from_win
for /f "tokens=1,2,3" %%a in (%MYDIRPATH_CODES_CONFIG%\_scp_to_remote.config) do (
	set USER=%%a
	set HOST=%%b
	set PASSWORD=%%c
)

if "%~1"=="" (
    echo ˆø”‚ğw’è‚µ‚Ä‚­‚¾‚³‚¢Bˆ—‚ğ’†’f‚µ‚Ü‚·B
    pause
    exit /b
)
set SRCPATH_WPATH=%~1
echo %SRCPATH_WPATH%
for /f "usebackq" %%A in (`wsl wslpath -a "%SRCPATH_WPATH%"`) do set SRCPATH_UPATH=%%A
echo %SRCPATH_UPATH%

wsl expect -c "spawn scp -r %SRCPATH_UPATH% %USER%@%HOST%:%DST_DIR_UPATH% ; expect password: ; send %PASSWORD%\r ; expect $ ; interact"
pause
