@echo off

:: usage: OpenCmdPromptAsRunas.bat [<work_dir_path>]

setlocal ENABLEDELAYEDEXPANSION
set LOG_FILE=%MYDIRPATH_DESKTOP%\OpenCmdPromptAsRunasArgs.log

whoami /priv | find "SeDebugPrivilege" > nul
if %errorlevel% neq 0 (
	echo %1> %LOG_FILE%
	@powershell start-process %~0 -verb runas
	exit
) else (
	set WORK_DIR=
	for /f %%i in (%LOG_FILE%) do (
		set WORK_DIR=%%i
	)
	del %LOG_FILE%
::	echo !WORK_DIR!
::	pause
	if "!WORK_DIR!" == "" (
		"%windir%\system32\cmd.exe" /k "title コマンドプロンプト"
	) else (
		"%windir%\system32\cmd.exe" /k "title コマンドプロンプト && cd !WORK_DIR! && cls"
	)
)

