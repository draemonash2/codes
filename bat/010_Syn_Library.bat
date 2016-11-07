@echo off
if {%MYPATH_CODE_BAT%} == {0} (
	echo target environment variable is nothing!
	pause
	exit /B 0
)
call %MYPATH_CODE_BAT%\lib\010_Def_Datetime.bat

set SRC=Z:\
set DSTBASE=X:\800_BackUp_Library

set DST=%DSTBASE%\data
set LOG=%DSTBASE%\log\%~n0_%datetime%.log

echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC%
echo ### Destination Path is %DST%
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR /XD "System Volume Information" >> %LOG%
echo ############### Finish! ##################
pause

echo. >> %LOG%
