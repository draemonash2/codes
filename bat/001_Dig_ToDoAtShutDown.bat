@echo off
call lib\010_Def_Datetime.bat

set LOGDIR=%~n0.log

echo Execute date is %date% %time% >> %LOGDIR%
echo {{{ >> %LOGDIR%

echo ######### Digest ToDoAtShutDown! #########
echo ####        Wait for a while ...      ####
echo %date% %time% >> %LOGDIR%
ruby ..\ruby\dig_ToDoAtShutDown.rb >> %LOGDIR%
echo ############### Finish! ##################
echo.

echo }}} >> %LOGDIR%
echo. >> %LOGDIR%
