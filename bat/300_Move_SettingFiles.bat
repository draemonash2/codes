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
	set EXEC_SCRIPT=%~dp0..\vbs\300_EvacuateSettingFiles.vbs
	echo.
	set /p ANS="�u�ޔ��v�ł�낵���ł����H [y/n] : "
	if !ANS! == y (
		rem
	) else (
		exit /B 0
	)
) else if %MOVE_TYPE% == r (
	set EXEC_SCRIPT=%~dp0..\vbs\301_RestoreSettingFiles.vbs
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

set DST_ROOT_PATH=C:\Users\draem_000\Documents\GoogleDrive\100_Programs
set LOGFILE_PATH=%DST_ROOT_PATH%\move_setting_files.log

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
::	"C:\Users\draem_000\Documents\GoogleDrive\Settings\PDF X-Change Viewer"
:: �ʃh���C�u����̃V���{���b�N�����N�͍쐬�s��
::	call "%EXEC_SCRIPT%"	"Z:\_ScratchLIVE_"																	"%DST_ROOT_PATH%\setting\Serato\_ScratchLIVE_"					"%LOGFILE_PATH%"
::	call "%EXEC_SCRIPT%"	"Z:\_ScratchLIVE_Backup"															"%DST_ROOT_PATH%\setting\Serato\_ScratchLIVE_Backup"			"%LOGFILE_PATH%"

:: #######################################################
:: ### �{����
:: #######################################################
	call "%EXEC_SCRIPT%"	"C:\prg_exe"																		"%DST_ROOT_PATH%\program\prg_exe"								"%LOGFILE_PATH%"
	
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Roaming\Scooter Software\Beyond Compare 3"				"%DST_ROOT_PATH%\setting\Beyond Compare 3\Beyond Compare 3"		"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Local\Kinza\User Data"									"%DST_ROOT_PATH%\setting\Kinza\User Data"						"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Roaming\Mozilla\Firefox"								"%DST_ROOT_PATH%\setting\Mozilla Firefox\Firefox"				"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Roaming\Hidemaruo\Hidemaru\Setting"						"%DST_ROOT_PATH%\setting\Hidemaru\Setting"						"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Roaming\foobar2000"										"%DST_ROOT_PATH%\setting\foobar2000\foobar2000"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Roaming\GRETECH\GomPlayer"								"%DST_ROOT_PATH%\setting\GomPlayer\GomPlayer"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Roaming\KeePass"										"%DST_ROOT_PATH%\setting\KeePass\KeePass"						"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Local\Thunderbird\Profiles"								"%DST_ROOT_PATH%\setting\Mozilla Thunderbird\Profiles"			"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Local\CherryPlayer\CherryPlayer 2.0"					"%DST_ROOT_PATH%\setting\CherryPlayer\CherryPlayer 2.0"			"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\AppData\Local\Icaros\IcarosCache"								"%DST_ROOT_PATH%\setting\Icaros\IcarosCache"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\Music\_Serato_"													"%DST_ROOT_PATH%\setting\Serato\_Serato_"						"%LOGFILE_PATH%"
		call "%EXEC_SCRIPT%"	"C:\Users\draem_000\Music\_Serato_Backup"										"%DST_ROOT_PATH%\setting\Serato\_Serato_Backup"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\Users\draem_000\Music\iTunes\iTunes Library Backup"								"%DST_ROOT_PATH%\setting\iTunes\iTunes Library Backup"			"%LOGFILE_PATH%"
		call "%EXEC_SCRIPT%"	"C:\Users\draem_000\Music\iTunes\iTunes Library Extras.itdb"					"%DST_ROOT_PATH%\setting\iTunes\iTunes Library Extras.itdb"		"%LOGFILE_PATH%"
		call "%EXEC_SCRIPT%"	"C:\Users\draem_000\Music\iTunes\iTunes Library Genius.itdb"					"%DST_ROOT_PATH%\setting\iTunes\iTunes Library Genius.itdb"		"%LOGFILE_PATH%"
		call "%EXEC_SCRIPT%"	"C:\Users\draem_000\Music\iTunes\iTunes Library.itl"							"%DST_ROOT_PATH%\setting\iTunes\iTunes Library.itl"				"%LOGFILE_PATH%"
		call "%EXEC_SCRIPT%"	"C:\Users\draem_000\Music\iTunes\iTunes Music Library.xml"						"%DST_ROOT_PATH%\setting\iTunes\iTunes Music Library.xml"		"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\prg\Everything\Everything.ini"													"%DST_ROOT_PATH%\setting\Everything\Everything.ini"				"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\prg\Honeyview\config.ini"														"%DST_ROOT_PATH%\setting\Honeyview\config.ini"					"%LOGFILE_PATH%"
	call "%EXEC_SCRIPT%"	"C:\prg\STEP038bin\SuperTagEditor.ini"												"%DST_ROOT_PATH%\setting\STEP038bin\SuperTagEditor.ini"			"%LOGFILE_PATH%"
	
	cmd.exe /c "%LOGFILE_PATH%"

endlocal
