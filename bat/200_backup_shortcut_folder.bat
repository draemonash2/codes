@echo off

:: backup shortcut folder
set ROOT_PATH=%CD%\..
set SRC_PATH=%ROOT_PATH%\shortcut
set DST_PATH=%ROOT_PATH%\shortcut_bak
set LOG_PATH=%~n0.log
set OPT=%OPT% /MIR
set OPT=%OPT% /R:5
set OPT=%OPT% /W:30
set OPT=%OPT% /SL
robocopy "%SRC_PATH%" "%DST_PATH%" %OPT% > "%LOG_PATH%"

echo.>> "%LOG_PATH%"

:: rename .lnk => .lnkbak
(
	FOR /R "%DST_PATH%" %%i IN (*.lnk) DO (
		echo %%i
		rename "%%i" "%%~ni.lnkbak"
	)
) >> "%LOG_PATH%"

pause
