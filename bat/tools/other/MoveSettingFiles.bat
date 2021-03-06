@echo off

setlocal ENABLEDELAYEDEXPANSION

whoami /PRIV | FIND "SeLoadDriverPrivilege" > NUL
if errorlevel 1 (
	echo 管理者権限で実行してください。
	pause
	exit /B 0
)
echo インストール済みソフトウェアの設定ファイルを移動します。
echo.
echo 実行前に以下の注意事項を読んでください。
echo    ・実行中のプログラムは閉じてください
echo    ・オンラインストレージのアクセスを停止してください。
echo.
set /p MOVE_TYPE="設定ファイルを退避する (e) か復帰するか (r) を選択してください。[e/r] : "
if %MOVE_TYPE% == e (
	set EXEC_SCRIPT_PATH=%~dp0..\vbs\300_EvacuateSettingFiles.vbs
	echo.
	set /p ANS="「退避」でよろしいですか？ [y/n] : "
	if !ANS! == y (
		rem
	) else (
		exit /B 0
	)
) else if %MOVE_TYPE% == r (
	set EXEC_SCRIPT_PATH=%~dp0..\vbs\301_RestoreSettingFiles.vbs
	echo.
	set /p ANS="「復帰」でよろしいですか？ [y/n] : "
	if !ANS! == y (
		rem
	) else (
		exit /B 0
	)
) else (
	echo e か r を選択してください。
	pause
	exit /B 0
)

set DST_ROOT_PATH=%USERPROFILE%\Documents\Amazon Drive\100_Programs
set LOGFILE_PATH=%CD%\100_move_setting_files.log
::set LOGFILE_PATH=%DST_ROOT_PATH%\_script\100_move_setting_files.log

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
:: ### 実施不可
:: #######################################################
:: セットアップの移行はインポート/エクスポートにて行うため、シンボリックリンクは作成しない。
::	"%USERPROFILE%\Documents\GoogleDrive\Settings\PDF X-Change Viewer"
:: 別ドライブからのシンボリックリンクは作成不可
::	call "%EXEC_SCRIPT_PATH%"		"Z:\_ScratchLIVE_"														"%DST_ROOT_PATH%\setting\Serato\_ScratchLIVE_"					"%LOGFILE_PATH%"
::	call "%EXEC_SCRIPT_PATH%"		"Z:\_ScratchLIVE_Backup"												"%DST_ROOT_PATH%\setting\Serato\_ScratchLIVE_Backup"			"%LOGFILE_PATH%"

:: #######################################################
:: ### 本処理
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
