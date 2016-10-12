@echo off
call lib\010_Def_Datetime.bat

set SRC=C:\Users\draem_000\Documents\Dropbox
set DSTBASE=X:\820_BackUp_Dropbox

set DST=%DSTBASE%\data
set LOG=%DSTBASE%\log\%~n0_%datetime%.log

echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC%
echo ### Destination Path is %DST%
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR /XD "System Volume Information" >> %LOG%
echo ############### Finish! ##################

echo. >> %LOG%
