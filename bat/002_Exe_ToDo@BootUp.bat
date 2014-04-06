@echo off
call lib\010_Def_Datetime.bat

set LOGDIR=.\log\%~n0_%datetime%.log

echo ########## Digest ToDoAtBootUP! ##########
echo ####        Wait for a while ...
cd C:\Users\TatsuyaEndo\Desktop
echo ############### Finish! ##################
echo.

echo. >> %LOGDIR%
