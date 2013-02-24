@echo off
call lib\010_Def_Datetime.bat
	
set LOGDIR=%~n0.log

echo Execute date is %date% %time% >> %LOGDIR%
echo {{{ >> %LOGDIR%

echo ############## Sync Drives! ##############
set /p SRC="### Source      Path [ex. D:\] : "
set /p DST="### Destination Path [ex. E:\] : "
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR >> %LOGDIR%
echo ############### Finish! ##################
pause

echo }}} >> %LOGDIR%
echo. >> %LOGDIR%
