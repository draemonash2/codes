@echo off

set LOGDIR=%~n0.log

echo ######### Digest ToDoAtShutDown! #########
echo ####        Wait for a while ...      ####
ruby ..\ruby\dig_ToDoAtShutDown.rb >> %LOGDIR%
echo ############### Finish! ##################
echo.
