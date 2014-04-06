@echo off
call lib\010_Def_Datetime.bat

set LOGDIR=.\log\%~n0_%datetime%.log

echo ######### Digest ToDoAtShutDown! #########
echo ####        Wait for a while ...
:: Sync Dropbox Directry
	call .\012_Syn_Dropbox.bat >> %LOGDIR%
:: Hidden System Files
	call ..\vbs\HiddenSystemFiles.vbs >> %LOGDIR%
:: Commit Setting Files
::	cd C:\prg					>> %LOGDIR%
::	git add -u .				>> %LOGDIR%
::	git commit -m "Auto Commit" >> %LOGDIR%
::	git push "setting"			>> %LOGDIR%
echo ############### Finish! ##################
echo.

echo. >> %LOGDIR%

timeout 5

:: シャットダウン
C:\WINDOWS\system32\shutdown -s -t 5
