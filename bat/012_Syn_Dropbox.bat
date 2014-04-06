:: ==============================================
:: <<実行時の注意>>
::   本バッチファイル実行前に、
::     WScript.Quit(WeekDay(Date))
::   と記述された VBS ファイルが C:\Windows 配下に
::   格納されていることを確認する。
::   存在しない場合は、格納しておく。
:: ==============================================

@echo off

cscript /b C:\Windows\CheckWeekDay.vbs

:: 月曜日のみ実行
if not %errorlevel% == 2 goto :END

	call lib\010_Def_Datetime.bat

	set LOGDIR=.\log\%~n0_%datetime%.log

	set SRC=C:\Users\TatsuyaEndo\Dropbox
	set DST=Z:\100_Documents\100_PC\90_BackUp\Dropbox

	echo ############### Sync Drive! ##############
	echo ### Source      Path is %SRC%
	echo ### Destination Path is %DST%
	set /p ANS="### Please press any key ..."
	echo ### Wait for a while ...
	echo {{{ >> %LOGDIR%
	robocopy %SRC% %DST% /MIR >> %LOGDIR%
	echo }}} >> %LOGDIR%
	echo ############### Finish! ##################
	pause

	echo. >> %LOGDIR%

:END
