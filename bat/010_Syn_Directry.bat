@echo off
if {%MYPATH_CODES%} == {0} (
	echo target environment variable is nothing!
	pause
	exit /B 0
)
call %MYPATH_CODES%\bat\lib\010_Def_Datetime.bat

set LOGDIR=%MYPATH_CODES%\bat\%~n0_%datetime%.log

echo ############# Sync Directry! #############
set /p SRC="### Source      Path [ex. D:\] : "
set /p DST="### Destination Path [ex. E:\] : "
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR >> %LOGDIR%
echo ############### Finish! ##################
pause

echo. >> %LOGDIR%
