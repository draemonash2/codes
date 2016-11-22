@echo off

set LOGDIR=%~dp0%~n0.log

echo ############# Sync Directry! #############
set /p SRC="### Source      Path [ex. D:\] : "
set /p DST="### Destination Path [ex. E:\] : "
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR >> %LOGDIR%
echo ############### Finish! ##################
pause

echo. >> %LOGDIR%
