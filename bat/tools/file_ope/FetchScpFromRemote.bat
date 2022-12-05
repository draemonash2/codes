@echo off
set tm=%time: =0%
set dt=%date:/=%
set DST_DIR_WPATH=%HOMEDRIVE%%HOMEPATH%\Desktop\_scp_from_remote_%dt:~2,6%-%tm:~0,2%%tm:~3,2%%tm:~6,2%
set SRC_DIR_UPATH=~/_scp_to_xxx
for /f "tokens=1,2,3" %%a in (%MYDIRPATH_CODES_CONFIG%\_ssh_target_a.config) do (
	set HOST=%%a
	set USER=%%b
	set PASSWORD=%%c
)

echo %DST_DIR_WPATH%
for /f "usebackq" %%A in (`wsl wslpath -a "%DST_DIR_WPATH%"`) do set DST_DIR_UPATH=%%A
echo %DST_DIR_UPATH%

wsl mkdir -p %DST_DIR_UPATH%
wsl expect -c "spawn scp -r %USER%@%HOST%:%SRC_DIR_UPATH%/* %DST_DIR_UPATH% ; expect password: ; send %PASSWORD%\r ; expect $ ; interact"
pause
