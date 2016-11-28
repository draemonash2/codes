@echo off

set SRC=C:\Users\draem_000\Documents\Dropbox
set DSTBASE=\\RASPBERRYPI\pockethdd\820_BackUp_Dropbox
set DST=%DSTBASE%\data
set LOG=%DSTBASE%\%~n0.log
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

echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC%
echo ### Destination Path is %DST%
echo ### Wait for a while ...
robocopy %SRC% %DST% %OPT% >NUL 2>&1
echo ############### Finish! ##################
