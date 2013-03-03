@echo off
call lib\010_Def_Datetime.bat

set LOGDIR=.\log\%~n0_%datetime%.log

echo ############# Sync Directry! #############
set /p SRC="### Source      Path [ex. D:\] : "
set /p DST="### Destination Path [ex. E:\] : "
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
echo {{{ >> %LOGDIR%
robocopy %SRC% %DST% /MIR >> %LOGDIR%
echo }}} >> %LOGDIR%
echo ############### Finish! ##################
pause

echo. >> %LOGDIR%
