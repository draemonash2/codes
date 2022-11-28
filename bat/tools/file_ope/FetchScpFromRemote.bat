@echo off
set DST_DIR_WPATH=%HOMEDRIVE%%HOMEPATH%\Desktop\_scp_from_remote
set SRC_DIR_UPATH=~/_scp_to_win
for /f "tokens=1,2,3" %%a in (%MYDIRPATH_CODES_CONFIG%\_scp_to_remote.config) do (
	set USER=%%a
	set HOST=%%b
	set PASSWORD=%%c
)

echo %DST_DIR_WPATH%
for /f "usebackq" %%A in (`wsl wslpath -a "%DST_DIR_WPATH%"`) do set DST_DIR_UPATH=%%A
echo %DST_DIR_UPATH%

wsl mkdir -p %DST_DIR_UPATH%
wsl expect -c "spawn scp -r %USER%@%HOST%:%SRC_DIR_UPATH%/* %DST_DIR_UPATH% ; expect password: ; send %PASSWORD%\r ; expect $ ; interact"
pause