@echo off

setlocal ENABLEDELAYEDEXPANSION

whoami /PRIV | FIND "SeLoadDriverPrivilege" > NUL
if errorlevel 1 (
	echo �Ǘ��Ҍ����Ŏ��s���Ă��������B
	pause
	exit /B 0
)
echo �C���X�g�[���ς݃\�t�g�E�F�A�̐ݒ�t�@�C�����ړ����܂��B
echo.
echo ���s�O�Ɉȉ��̒��ӎ�����ǂ�ł��������B
echo    �E���s���̃v���O�����͕��Ă�������
echo    �E�I�����C���X�g���[�W�̃A�N�Z�X���~���Ă��������B
echo.
set /p MOVE_TYPE="�ݒ�t�@�C����ޔ����� (e) �����A���邩 (r) ��I�����Ă��������B[e/r] : "
if %MOVE_TYPE% == e (
	set EXEC_SCRIPT_PATH=%~dp0..\vbs\300_EvacuateSettingFiles.vbs
	echo.
	set /p ANS="�u�ޔ��v�ł�낵���ł����H [y/n] : "
	if !ANS! == y (
		rem
	) else (
		exit /B 0
	)
) else if %MOVE_TYPE% == r (
	set EXEC_SCRIPT_PATH=%~dp0..\vbs\301_RestoreSettingFiles.vbs
	echo.
	set /p ANS="�u���A�v�ł�낵���ł����H [y/n] : "
	if !ANS! == y (
		rem
	) else (
		exit /B 0
	)
) else (
	echo e �� r ��I�����Ă��������B
	pause
	exit /B 0
)

set DST_ROOT_PATH=%USERPROFILE%\Documents\Amazon Drive\100_Programs
set LOGFILE_PATH=%DST_ROOT_PATH%\_script\move_setting_files.log

echo.>> "%LOGFILE_PATH%"
echo _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/ >> "%LOGFILE_PATH%"
if %MOVE_TYPE% == e (
	echo _/ start evacuate setting files                                               _/ >> "%LOGFILE_PATH%"
) else (
	echo _/ start restore setting files                                                _/ >> "%LOGFILE_PATH%"
)
echo _/ time is %date% %time%                                             _/ >> "%LOGFILE_PATH%"
echo _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/ >> "%LOGFILE_PATH%"

:: #######################################################
:: ### ���{�s��
:: #######################################################
:: �Z�b�g�A�b�v�̈ڍs�̓C���|�[�g/�G�N�X�|�[�g�ɂčs�����߁A�V���{���b�N�����N�͍쐬���Ȃ��B
::	"%USERPROFILE%\Documents\GoogleDrive\Settings\PDF X-Change Viewer"
:: �ʃh���C�u����̃V���{���b�N�����N�͍쐬�s��
::	call "%EXEC_SCRIPT_PATH%"		"Z:\_ScratchLIVE_"														"%DST_ROOT_PATH%\setting\Serato\_ScratchLIVE_"					"%LOGFILE_PATH%"
::	call "%EXEC_SCRIPT_PATH%"		"Z:\_ScratchLIVE_Backup"												"%DST_ROOT_PATH%\setting\Serato\_ScratchLIVE_Backup"			"%LOGFILE_PATH%"

:: #######################################################
:: ### �{����
:: #######################################################
	:: setting
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Local\Kinza\User Data"							"%DST_ROOT_PATH%\setting\Kinza\User Data"						"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Local\HNXgrep"									"%DST_ROOT_PATH%\setting\HNXgrep\HNXgrep"						"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Local\Icaros"									"%DST_ROOT_PATH%\setting\Icaros\Icaros"							"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Roaming\GZ20"									"%DST_ROOT_PATH%\setting\EasyShot\GZ20"							"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Roaming\KT Software"								"%DST_ROOT_PATH%\setting\DeInput\KT Software"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Roaming\Mp3tag"									"%DST_ROOT_PATH%\setting\MP3Tag\MP3Tag"							"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Roaming\Team Hasebe"								"%DST_ROOT_PATH%\setting\TVClock\Team Hasebe"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Roaming\Audacity"								"%DST_ROOT_PATH%\setting\Audacity\Audacity"						"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Roaming\Subversion"								"%DST_ROOT_PATH%\setting\Subversion\Subversion"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Roaming\TortoiseGit"								"%DST_ROOT_PATH%\setting\TortoiseGit\TortoiseGit"				"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\AppData\Roaming\TortoiseSVN"								"%DST_ROOT_PATH%\setting\TortoiseSVN\TortoiseSVN"				"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\Music\_Serato_"											"%DST_ROOT_PATH%\setting\Serato\_Serato_"						"%LOGFILE_PATH%"
		call "%EXEC_SCRIPT_PATH%"	"%USERPROFILE%\Music\_Serato_Backup"									"%DST_ROOT_PATH%\setting\Serato\_Serato_Backup"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT_PATH%"		"%USERPROFILE%\Music\iTunes"											"%DST_ROOT_PATH%\setting\iTunes\iTunes"							"%LOGFILE_PATH%"
	
	:: program
	call "%EXEC_SCRIPT_PATH%"		"C:\prg_exe"															"%DST_ROOT_PATH%\program\prg_exe"								"%LOGFILE_PATH%"
	
	cmd.exe /c "%LOGFILE_PATH%"

endlocal
