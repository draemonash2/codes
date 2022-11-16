@echo off
set DST_DIR_WPATH=%HOMEDRIVE%%HOMEPATH%\Desktop\_scp_from_remote
set SRC_DIR_UPATH=~/_scp_to_win
set USER=user
set HOST=XXX.XXX.XXX.XXX
set PASSWORD=password

echo %DST_DIR_WPATH%
for /f "usebackq" %%A in (`wsl wslpath -a "%DST_DIR_WPATH%"`) do set DST_DIR_UPATH=%%A
echo %DST_DIR_UPATH%

wsl mkdir -p %DST_DIR_UPATH%
wsl expect -c "spawn scp -r %USER%@%HOST%:%SRC_DIR_UPATH%/* %DST_DIR_UPATH% ; expect password: ; send %PASSWORD%\r ; expect $ ; interact"
pause
