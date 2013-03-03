@echo off
call lib\010_Def_Datetime.bat

set LOGDIR=.\log\%~n0_%datetime%.log

set SRC=D:\
set DST=E:\

echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC%
echo ### Destination Path is %DST%
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
echo {{{ >> %LOGDIR%
robocopy %SRC% %DST% /MIR >> %LOGDIR%
echo }}} >> %LOGDIR%
echo ############### Finish! ##################
pause

echo. >> %LOGDIR%
