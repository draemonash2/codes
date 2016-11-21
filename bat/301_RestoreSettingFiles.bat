@echo off

::<<�T�v>>
::	�i�[���ύX�����v���O�����̐ݒ�t�@�C�����A���̊i�[��ɖ߂��B
::	���̍ہA�쐬�����V���{���b�N�����N���폜���āA�i�[��ύX�O��
::	��Ԃɕ�������B
::
::<<����>>
::	�����P�F�ޔ����t�@�C��/�t�H���_�p�X
::	�����Q�F�ޔ��t�@�C��/�t�H���_�p�X
::
::<<������>>
::	�P�D�V���{���b�N�����N�폜
::	�Q�D�ޔ����t�H���_�ւ̃V���[�g�J�b�g�폜
::	�R�D�t�@�C��/�t�H���_�ړ��i�ޔ��ˑޔ����j
::	�S�D�ޔ����ɍ쐬�����t�H���_���폜
::
::<<�o��>>
::	�E�ޔ��t�@�C��/�t�H���_�p�X�����݂��Ȃ��ꍇ�A�������Ȃ��B
::	�E�w�肷��p�X�̓t�@�C��/�t�H���_�ǂ���ł��B
::	�E�Ǘ��Ҍ����Ŏ��s���邱�ƁB�Ǘ��Ҍ����łȂ��ꍇ�A�I������B

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
