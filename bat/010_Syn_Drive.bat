@echo off
call lib\010_Def_Datetime.bat

set LOGDIR=D:\%~n0_%datetime%.log

set SRC=Z:\
set DST=D:\BackupFile

echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC%
echo ### Destination Path is %DST%
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR /XD "System Volume Information" >> %LOGDIR%
echo ############### Finish! ##################
pause

echo. >> %LOGDIR%
