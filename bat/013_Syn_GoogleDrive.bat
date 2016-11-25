@echo off

set SRC=C:\Users\draem_000\Documents\GoogleDrive
set DSTBASE=\\RASPBERRYPI\pockethdd\830_BackUp_GoogleDrive

set DST=%DSTBASE%\data
set LOG=%DSTBASE%\%~n0.log

echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC%
echo ### Destination Path is %DST%
set /p ANS="### Please press any key ..."
echo ### Wait for a while ...
robocopy %SRC% %DST% /MIR /XD "System Volume Information" > %LOG%
echo ############### Finish! ##################
pause

echo. >> %LOG%
