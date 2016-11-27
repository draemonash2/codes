@echo off

:: 指定フォルダパスに含まれるフォルダが空か判定し、空フォルダなら削除する。
:: 引数未指定もしくは指定パスが存在しない場合、FALSE を出力。
:: それ以外の場合、TRUE を出力。
:: 出力したくない場合、「call <本バッチファイル> >NUL 2>&1」として呼び出す。

setlocal

if "%~1" == "" goto negative

set ARG=%~1
::	echo argument : %ARG%

::末尾に \ が付与されている場合、\ を削除してやり直し
if %ARG:~-1% == \ (
	call %0 "%ARG:~0,-1%"
)

set TRGT_DIR_PATH_TMP=%~1
if %TRGT_DIR_PATH_TMP:~-1% == \ (
	set TRGT_DIR_PATH=%TRGT_DIR_PATH_TMP:~0,-1%
) else (
	set TRGT_DIR_PATH=%TRGT_DIR_PATH_TMP%
)
::	echo target dir     : %TRGT_DIR_PATH%
if not exist "%TRGT_DIR_PATH%" goto negative

set TRGT_PAR_DIR_PATH_TMP=%~dp1
set TRGT_PAR_DIR_PATH=%TRGT_PAR_DIR_PATH_TMP:~0,-1%
::	echo target par dir : %TRGT_PAR_DIR_PATH%

:positive
set FILE_EXISTS=FALSE
for    %%i in ("%TRGT_DIR_PATH%\*") do set FILE_EXISTS=TRUE
for /d %%i in ("%TRGT_DIR_PATH%\*") do set FILE_EXISTS=TRUE
if %FILE_EXISTS% == FALSE (
	rmdir /s /q "%TRGT_DIR_PATH%" >NUL 2>&1
	call %0 "%TRGT_PAR_DIR_PATH%"
) else (
	echo TRUE
)
goto end

:negative
echo FALSE

:end
endlocal
