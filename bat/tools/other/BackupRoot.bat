@echo off

setlocal ENABLEDELAYEDEXPANSION

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: 管理者権限チェック
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
whoami /PRIV | FIND "SeLoadDriverPrivilege" > NUL
if errorlevel 1 (
	echo [error] please execute on runas mode!
	pause
	exit /B 0
)

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: robocopy オプション判定
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
set SRC_PATH=C:\Users\draem\_root
set DST_BASE_PATH=X:\810_BackUp_PC\latest
set DST_PATH=%DST_BASE_PATH%\root
set OPT=
set OPT=!OPT! /MIR
set OPT=!OPT! /R:5
set OPT=!OPT! /W:30
set OPT=!OPT! /SL
set OPT=!OPT! /XD "System Volume Information"
set IDX_MAX=1

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: アクティブディレクトリ判定
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
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
:: 同期実行
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
echo ############### Sync Drive! ##############
echo ### Source      Path is %SRC_PATH%
echo ### Destination Path is %DST_PATH%
echo ### Wait for a while ...
robocopy "%SRC_PATH%" "%DST_PATH%" %OPT% > "%LOG_PATH%"
::echo robocopy "%SRC_PATH%" "%DST_PATH%" %OPT% "%LOG_PATH%"
echo ############### Finish! ##################

endlocal
