@echo off

::<<�T�v>>
::  �v���O�����̐ݒ�t�@�C���̊i�[���ύX����B
::  �Ȃ��A���i�[�悩��V�i�[��Ɍ����ăV���{���b�N�����N���쐬���邽�߁A
::  �v���O�������ł̐ݒ�͕s�v�B
::
::<<����>>
::  �����P�F�ޔ����t�@�C��/�t�H���_�p�X
::  �����Q�F�ޔ��t�@�C��/�t�H���_�p�X
::
::<<������>>
::  �P�D�ޔ��t�@�C��/�t�H���_�폜
::  �Q�D�ޔ��t�@�C��/�t�H���_�쐬
::  �R�D�t�@�C��/�t�H���_�ړ��i�ޔ����ˑޔ��j
::  �S�D�ޔ����ˑޔ��ւ̃V���{���b�N�����N�쐬
::  �T�D�ޔ����t�H���_�ւ̃V���[�g�J�b�g���쐬
::
::<<�o��>>
::  �E���łɃV���{���b�N�����N���쐬����Ă���ꍇ�́A�������Ȃ��B
::  �E�ޔ����t�@�C��/�t�H���_�p�X�����݂��Ȃ��ꍇ�A�������Ȃ��B
::  �E�w�肷��p�X�̓t�@�C��/�t�H���_�ǂ���ł��B
::  �E�Ǘ��Ҍ����Ŏ��s���邱�ƁB�Ǘ��Ҍ����łȂ��ꍇ�A�I������B

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
