@echo off

set LOG=%~dp0%~n0.log
set OPT=
set OPT=%OPT% /MIR
set OPT=%OPT% /SL
set OPT=%OPT% /XD "System Volume Information"
set OPT=%OPT% /LOG:%LOG%

whoami /PRIV | FIND "SeLoadDriverPrivilege" > NUL
if errorlevel 1 (
	echo ### result : [error  ] please execute on runas mode!
	pause
	exit /B 0
)

echo ############# Sync Directry! #############
set /p SRC="### Source      Path [ex. D:\] : "
set /p DST="### Destination Path [ex. E:\] : "
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
robocopy %SRC% %DST% %OPT% >NUL 2>&1
echo ############### Finish! ##################
pause
