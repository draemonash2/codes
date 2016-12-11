@echo off
::第一引数：対象
::  /l : Library
::  /d : Dropbox
::  /g : GoogleDrive
::  /a : AmazonDrive
::
::第二引数：実行モード
::  /close   : 処理終了後、コンソールを閉じる
::  /suspend : 処理終了後、コンソールを開いたままにする
::  指定なし : 処理終了後、コンソールを開いたままにする
::
::第三引数：多重化数

setlocal ENABLEDELAYEDEXPANSION

whoami /PRIV | FIND "SeLoadDriverPrivilege" > NUL
if errorlevel 1 (
	echo [error] please execute on runas mode!
	pause
	exit /B 0
)

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

set OPT=
if %1 == /l (
	set SRC_PATH=Z:\
	set DST_BASE_PATH=\\RASPBERRYPI\pockethdd\800_BackUp_Library
	set OPT=!OPT! /MIR
	set OPT=!OPT! /SL
	set OPT=!OPT! /XD "System Volume Information"
) else if %1 == /d (
	set SRC_PATH=C:\Users\draem_000\Documents\Dropbox
	set DST_BASE_PATH=\\RASPBERRYPI\pockethdd\820_BackUp_Dropbox
	set OPT=!OPT! /MIR
	set OPT=!OPT! /SL
	set OPT=!OPT! /XD "System Volume Information"
) else if %1 == /g (
	set SRC_PATH=C:\Users\draem_000\Documents\GoogleDrive
	set DST_BASE_PATH=\\RASPBERRYPI\pockethdd\830_BackUp_GoogleDrive
	set OPT=!OPT! /MIR
	set OPT=!OPT! /SL
	set OPT=!OPT! /XF "Current Session"
	set OPT=!OPT! /XF "Current Tabs"
) else if %1 == /a (
	set SRC_PATH=C:\Users\draem_000\Documents\Amazon Drive
	set DST_BASE_PATH=\\RASPBERRYPI\pockethdd\840_BackUp_AmazonDrive
	set OPT=!OPT! /MIR
	set OPT=!OPT! /SL
	set OPT=!OPT! /XF "Current Session"
	set OPT=!OPT! /XF "Current Tabs"
) else (
	echo [error] argument1 error! arg1 is %1
	pause
	exit /B 0
)
set DST_PATH=%DST_BASE_PATH%\data

if "%~2" == "/close" (
	set EXEC_MODE=CLOSE
) else if "%~2" == "/suspend" (
	set EXEC_MODE=SUSPEND
) else (
	echo [error] argument2 error! arg2 is %2
	pause
	exit /B 0
)

set IDX_MAX=%~3
set PREV_ACTIVE_DIR_NUM=1
set CURR_ACTIVE_DIR_NUM=1
for /l %%i in (1,1,%IDX_MAX%) do (
	if exist %DST_PATH%_%%i_is_active_directory (
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
del %DST_PATH%_%PREV_ACTIVE_DIR_NUM%_is_active_directory >NUL 2>&1
echo.> %DST_PATH%_%CURR_ACTIVE_DIR_NUM%_is_active_directory
set DST_PATH=%DST_PATH%_%CURR_ACTIVE_DIR_NUM%
set LOG_PATH=%DST_BASE_PATH%\%~n0_%CURR_ACTIVE_DIR_NUM%.log

echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC_PATH%
echo ### Destination Path is %DST_PATH%
if "%EXEC_MODE%" == "SUSPEND" (
	set /p ANS="### Please press any key ..."
) else (
	rem
)
echo ### Wait for a while ...
robocopy "%SRC_PATH%" "%DST_PATH%" %OPT% /LOG:%LOG_PATH% >NUL 2>&1
echo ############### Finish! ##################
if "%EXEC_MODE%" == "SUSPEND" (
	pause
) else (
	rem
)

endlocal
