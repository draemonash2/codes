@echo off
::�������F�Ώ�
::  /l : Library
::  /d : Dropbox
::  /g : GoogleDrive
::  /a : AmazonDrive
::
::�������F���s���[�h
::  /close   : �����I����A�R���\�[�������
::  /suspend : �����I����A�R���\�[�����J�����܂܂ɂ���
::  �w��Ȃ� : �����I����A�R���\�[�����J�����܂܂ɂ���
::
::��O�����F���d����

setlocal ENABLEDELAYEDEXPANSION

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: �ݒ�l
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	set DST_BASE=\\RASPBERRYPI\pockethdd
::	set DST_BASE=D:

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: �Ǘ��Ҍ����`�F�b�N
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
whoami /PRIV | FIND "SeLoadDriverPrivilege" > NUL
if errorlevel 1 (
	echo [error] please execute on runas mode!
	pause
	exit /B 0
)

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: �����`�F�b�N
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                 set ARG_NUM=9
if "%~9" == "" ( set ARG_NUM=8 ) else ( goto exec )
if "%~8" == "" ( set ARG_NUM=7 ) else ( goto exec )
if "%~7" == "" ( set ARG_NUM=6 ) else ( goto exec )
if "%~6" == "" ( set ARG_NUM=5 ) else ( goto exec )
if "%~5" == "" ( set ARG_NUM=4 ) else ( goto exec )
if "%~4" == "" ( set ARG_NUM=3 ) else ( goto exec )
if "%~3" == "" ( set ARG_NUM=2 ) else ( goto exec )
if "%~2" == "" ( set ARG_NUM=1 ) else ( goto exec )
if "%~1" == "" ( set ARG_NUM=0 ) else ( goto exec )
:exec
if %ARG_NUM% == 3 (
	rem
) else (
	echo [error] argument number error! argument number is %ARG_NUM%
	pause
	exit /B 0
)

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: robocopy �I�v�V��������
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
set OPT=
if %1 == /l (
	set SRC_PATH=Z:
	set DST_BASE_PATH=!DST_BASE!\821_BackUp_Library
	set OPT=!OPT! /MIR
	set OPT=!OPT! /R:5
	set OPT=!OPT! /W:30
	set OPT=!OPT! /SL
	set OPT=!OPT! /XD "System Volume Information"
) else if %1 == /d (
	set SRC_PATH=C:\Users\draem_000\Documents\Dropbox
	set DST_BASE_PATH=!DST_BASE!\822_BackUp_Dropbox
	set OPT=!OPT! /MIR
	set OPT=!OPT! /R:5
	set OPT=!OPT! /W:30
	set OPT=!OPT! /SL
	set OPT=!OPT! /XD "System Volume Information"
) else if %1 == /a (
	set SRC_PATH=C:\Users\draem_000\Documents\Amazon Drive
	set DST_BASE_PATH=!DST_BASE!\823_BackUp_AmazonDrive
	set OPT=!OPT! /MIR
	set OPT=!OPT! /R:5
	set OPT=!OPT! /W:30
	set OPT=!OPT! /SL
	set OPT=!OPT! /XF "Current Session"
	set OPT=!OPT! /XF "Current Tabs"
) else (
	echo [error] argument1 error! arg1 is %1
	pause
	exit /B 0
)
set DST_PATH=%DST_BASE_PATH%\data

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: ���s���[�h�i�I�� or ��~�j����
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
if "%~2" == "/close" (
	set EXEC_MODE=CLOSE
) else if "%~2" == "/suspend" (
	set EXEC_MODE=SUSPEND
) else (
	echo [error] argument2 error! arg2 is %2
	pause
	exit /B 0
)

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: �A�N�e�B�u�f�B���N�g������
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
set IDX_MAX=%~3
set PREV_ACTIVE_DIR_NUM=1
set CURR_ACTIVE_DIR_NUM=1
for /l %%i in (1,1,%IDX_MAX%) do (
	if exist "%DST_PATH%_%%i is active directory" (
		set PREV_ACTIVE_DIR_NUM=%%i
		if %%i==%IDX_MAX% (
			set CURR_ACTIVE_DIR_NUM=1
		) else (
			set /a "CURR_ACTIVE_DIR_NUM = %%i + 1"
		)
		goto break
	)
)
:break
del "%DST_PATH%_%PREV_ACTIVE_DIR_NUM% is active directory" >NUL 2>&1
echo.> "%DST_PATH%_%CURR_ACTIVE_DIR_NUM% is active directory"
set DST_PATH=%DST_PATH%_%CURR_ACTIVE_DIR_NUM%
set LOG_PATH=%DST_BASE_PATH%\%~n0_%CURR_ACTIVE_DIR_NUM%.log

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: �������s
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC_PATH%
echo ### Destination Path is %DST_PATH%
if "%EXEC_MODE%" == "SUSPEND" (
	set /p ANS="### Please press any key ..."
) else (
	rem
)
echo ### Wait for a while ...
robocopy "%SRC_PATH%" "%DST_PATH%" %OPT% > "%LOG_PATH%"
::echo robocopy "%SRC_PATH%" "%DST_PATH%" %OPT% "%LOG_PATH%"
echo ############### Finish! ##################
if "%EXEC_MODE%" == "SUSPEND" (
	pause
) else (
	rem
)

endlocal
