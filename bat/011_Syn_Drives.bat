@echo off

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

if %ARG_NUM% == 2 (
	if "%~2" == "/close" (
		set EXEC_MODE=CLOSE
	) else if "%~2" == "/suspend" (
		set EXEC_MODE=SUSPEND
	) else (
		echo [error] argument2 error! arg2 is %2
		pause
		exit /B 0
	)
) else if %ARG_NUM% == 1 (
	set EXEC_MODE=SUSPEND
) else (
	echo [error] argument number error! argument number is %ARG_NUM%
	pause
	exit /B 0
)

set OPT=
if %1 == /l (
	set SRC=Z:\
	set DSTBASE=\\RASPBERRYPI\pockethdd\800_BackUp_Library
	set DST=!DSTBASE!\data
	set LOG=!DSTBASE!\%~n0.log
	set OPT=!OPT! /MIR
	set OPT=!OPT! /SL
	set OPT=!OPT! /XD "System Volume Information"
	set OPT=!OPT! /LOG:!LOG!
) else if %1 == /d (
	set SRC=C:\Users\draem_000\Documents\Dropbox
	set DSTBASE=\\RASPBERRYPI\pockethdd\820_BackUp_Dropbox
	set DST=!DSTBASE!\data
	set LOG=!DSTBASE!\%~n0.log
	set OPT=!OPT! /MIR
	set OPT=!OPT! /SL
	set OPT=!OPT! /XD "System Volume Information"
	set OPT=!OPT! /LOG:!LOG!
) else if %1 == /g (
	set SRC=C:\Users\draem_000\Documents\GoogleDrive
	set DSTBASE=\\RASPBERRYPI\pockethdd\830_BackUp_GoogleDrive
	set DST=!DSTBASE!\data
	set LOG=!DSTBASE!\%~n0.log
	set OPT=!OPT! /MIR
	set OPT=!OPT! /SL
	set OPT=!OPT! /XF "Current Session"
	set OPT=!OPT! /XF "Current Tabs"
	set OPT=!OPT! /LOG:!LOG!
) else if %1 == /a (
	set SRC=C:\Users\draem_000\Documents\Amazon Drive
	set DSTBASE=\\RASPBERRYPI\pockethdd\840_BackUp_AmazonDrive
	set DST=!DSTBASE!\data
	set LOG=!DSTBASE!\%~n0.log
	set OPT=!OPT! /MIR
	set OPT=!OPT! /SL
	set OPT=!OPT! /XF "Current Session"
	set OPT=!OPT! /XF "Current Tabs"
	set OPT=!OPT! /LOG:!LOG!
) else (
	echo [error] argument1 error! arg1 is %1
	pause
	exit /B 0
)

echo ############### Sync Drive! ##############
echo ### Source	   Path is %SRC%
echo ### Destination Path is %DST%
if "%EXEC_MODE%" == "SUSPEND" (
	set /p ANS="### Please press any key ..."
) else (
	rem
)
echo ### Wait for a while ...
robocopy "%SRC%" "%DST%" %OPT% >NUL 2>&1
echo ############### Finish! ##################
if "%EXEC_MODE%" == "SUSPEND" (
	pause
) else (
	rem
)

endlocal
