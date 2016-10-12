@echo off
call lib\010_Def_Datetime.bat

set SRC=C:\Users\draem_000\Documents\Dropbox
set DSTBASE=X:\120_BackUp_Dropbox

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
