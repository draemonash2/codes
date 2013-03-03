@echo off
call lib\010_Def_Datetime.bat

set LOGDIR=.\log\%~n0_%datetime%.log

echo ######### Digest ToDoAtShutDown! #########
echo ####        Wait for a while ...
echo {{{ >> %LOGDIR%
ruby ..\ruby\dig_ToDoAtShutDown.rb >> %LOGDIR%
echo }}} >> %LOGDIR%
echo ############### Finish! ##################
echo.

echo. >> %LOGDIR%
