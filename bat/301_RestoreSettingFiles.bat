@echo off

::<<概要>>
::	格納先を変更したプログラムの設定ファイルを、元の格納先に戻す。
::	その際、作成したシンボリックリンクを削除して、格納先変更前の
::	状態に復元する。
::
::<<引数>>
::	引数１：退避元ファイル/フォルダパス
::	引数２：退避先ファイル/フォルダパス
::
::<<処理順>>
::	１．シンボリックリンク削除
::	２．退避元フォルダへのショートカット削除
::	３．ファイル/フォルダ移動（退避先⇒退避元）
::	４．退避時に作成したフォルダを削除
::
::<<覚書>>
::	・退避先ファイル/フォルダパスが存在しない場合、処理しない。
::	・指定するパスはファイル/フォルダどちらでも可。
::	・管理者権限で実行すること。管理者権限でない場合、終了する。

setlocal

if {%MYPATH_CODES%} == {0} (
	echo target environment variable is nothing!
	pause
	goto end
)
set DERETE_EMPTY_FOLDERS_BAT=%MYPATH_CODES%\bat\lib\DeleteEmptyFolders.bat

echo #########################################################
echo ### src	: %~1
echo ### dst	: %~2

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

if not exist %DST_PATH% (
	echo ### result : [error  ] destination file/folder is missing!
	goto end
)

set SRC_ATTR=%~a1
set DST_ATTR=%~a2

if %DST_ATTR:~0,1%==d (
	echo ### target : folder
	if exist %SRC_PATH% rmdir /s /q %SRC_PATH% >NUL 2>&1
	if exist %SHORTCUT_PATH% rmdir /s /q %SHORTCUT_PATH% >NUL 2>&1
	move %DST_PATH% %SRC_PATH% >NUL 2>&1
	Call %DERETE_EMPTY_FOLDERS_BAT% %DST_PAR_PATH% >NUL 2>&1
	echo ### result : [success] setting files are restored!
) else (
	echo ### target : file
	if exist %SRC_PATH% del /a /q %SRC_PATH% >NUL 2>&1
	if exist %SHORTCUT_PATH% rmdir /s /q %SHORTCUT_PATH% >NUL 2>&1
	move %DST_PATH% %SRC_PATH% >NUL 2>&1
	Call %DERETE_EMPTY_FOLDERS_BAT% %DST_PAR_PATH% >NUL 2>&1
	echo ### result : [success] setting files are restored!
)

:end
endlocal
