@echo off

::<<概要>>
::  プログラムの設定ファイルの格納先を変更する。
::  なお、元格納先から新格納先に向けてシンボリックリンクを作成するため、
::  プログラム側での設定は不要。
::
::<<引数>>
::  引数１：退避元ファイル/フォルダパス
::  引数２：退避先ファイル/フォルダパス
::
::<<処理順>>
::  １．退避先ファイル/フォルダ削除
::  ２．退避先ファイル/フォルダ作成
::  ３．ファイル/フォルダ移動（退避元⇒退避先）
::  ４．退避元⇒退避先へのシンボリックリンク作成
::  ５．退避元フォルダへのショートカットを作成
::
::<<覚書>>
::  ・すでにシンボリックリンクが作成されている場合は、処理しない。
::  ・退避元ファイル/フォルダパスが存在しない場合、処理しない。
::  ・指定するパスはファイル/フォルダどちらでも可。
::  ・管理者権限で実行すること。管理者権限でない場合、終了する。

setlocal

echo #########################################################
echo ### src    : %~1
echo ### dst    : %~2

whoami /PRIV | FIND "SeLoadDriverPrivilege" > NUL
if errorlevel 1 (
	echo ### result : [error  ] please execute on runas mode!
	goto end
)

set IS_ERROR=FALSE
if "%~1" == "" set IS_ERROR=TRUE
if "%~2" == "" set IS_ERROR=TRUE
if %IS_ERROR%==TRUE (
	echo ### result : [error  ] argument number error!
	goto end
)

set SRC_PATH="%~1"
set DST_PATH="%~2"
set SRC_PAR_PATH="%~dp1"
set DST_PAR_PATH="%~dp2"
set SRC_NAME=%~n1%~x1
set DST_NAME=%~n2%~x2
set SHORTCUT_PATH="%~dp2%SRC_NAME%_linksrc"

::echo SRC_PATH			: %SRC_PATH%
::echo DST_PATH			: %DST_PATH%
::echo SRC_PAR_PATH		: %SRC_PAR_PATH%
::echo DST_PAR_PATH		: %DST_PAR_PATH%
::echo SRC_NAME			: %SRC_NAME%
::echo DST_NAME			: %DST_NAME%
::echo SHORTCUT_PATH	: %SHORTCUT_PATH%

if not exist %SRC_PATH% (
	echo ### result : [error  ] source file/folder is missing!
	goto end
)

set SRC_ATTR=%~a1
set DST_ATTR=%~a2

if %SRC_ATTR:~0,1%==d (
	echo ### target : folder
	if %SRC_ATTR:~8,1%==l (
		echo ### result : [error  ] setting files are already evacuated!
		goto end
	)
	if exist %DST_PATH% rmdir /s /q %DST_PATH% >NUL 2>&1
	mkdir %DST_PAR_PATH% >NUL 2>&1
	move %SRC_PATH% %DST_PATH% >NUL 2>&1
	mklink /d %SRC_PATH% %DST_PATH% >NUL 2>&1
	mklink /d %SHORTCUT_PATH% %SRC_PAR_PATH% >NUL 2>&1
	echo ### result : [success] setting files are evacuated!
) else (
	echo ### target : file
	if %SRC_ATTR:~8,1%==l (
		echo ### result : [error  ] setting files are already evacuated!
		goto end
	)
	if exist %DST_PATH% del /a /q %DST_PATH% >NUL 2>&1
	mkdir %DST_PAR_PATH% >NUL 2>&1
	move %SRC_PATH% %DST_PATH% >NUL 2>&1
	mklink %SRC_PATH% %DST_PATH% >NUL 2>&1
	mklink /d %SHORTCUT_PATH% %SRC_PAR_PATH% >NUL 2>&1
	echo ### result : [success] setting files are evacuated!
)

:end
endlocal
