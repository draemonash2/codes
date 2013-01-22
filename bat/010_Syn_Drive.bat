@echo off

set LOGDIR=%~n0.log

echo ############## Sync Drives! ##############
set /p SRC="### Source      Path [ex. D:\] : "
set /p DST="### Destination Path [ex. E:\] : "
pause
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR >> %LOGDIR%
echo ############### Finish! ##################
pause
