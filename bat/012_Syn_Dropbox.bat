:: ==============================================
:: <<���s���̒���>>
::   �{�o�b�`�t�@�C�����s�O�ɁA
::     WScript.Quit(WeekDay(Date))
::   �ƋL�q���ꂽ VBS �t�@�C���� C:\Windows �z����
::   �i�[����Ă��邱�Ƃ��m�F����B
::   ���݂��Ȃ��ꍇ�́A�i�[���Ă����B
:: ==============================================

@echo off

cscript /b C:\Windows\CheckWeekDay.vbs

:: ���j���̂ݎ��s
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
