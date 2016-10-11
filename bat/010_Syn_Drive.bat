@echo off
call lib\010_Def_Datetime.bat

set LOGDIR=X:\101_BackupFile_log\%~n0_%datetime%.log

set SRC=Z:\
set DST=X:\100_BackupFile

echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC%
echo ### Destination Path is %DST%
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR /XD "System Volume Information" >> %LOGDIR%
echo ############### Finish! ##################
pause

echo. >> %LOGDIR%
