@echo off

:: �w��t�H���_�p�X�Ɋ܂܂��t�H���_���󂩔��肵�A��t�H���_�Ȃ�폜����B
:: �������w��������͎w��p�X�����݂��Ȃ��ꍇ�AFALSE ���o�́B
:: ����ȊO�̏ꍇ�ATRUE ���o�́B
:: �o�͂������Ȃ��ꍇ�A�ucall <�{�o�b�`�t�@�C��> >NUL 2>&1�v�Ƃ��ČĂяo���B

setlocal

if "%~1" == "" goto negative

set ARG=%~1
::	echo argument : %ARG%

::������ \ ���t�^����Ă���ꍇ�A\ ���폜���Ă�蒼��
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
