@echo off
set DST_PATH=%CD%\..\shortcut_bak
set LOG_PATH=%CD%\%~n0.log
(
	FOR /R "%DST_PATH%" %%i IN (*.lnk) DO (
		echo %%i
		rename "%%i" "%%~ni.lnkbak"
	)
) > "%LOG_PATH%"
pause
