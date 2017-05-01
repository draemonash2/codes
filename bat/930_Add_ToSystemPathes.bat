@echo off
::	<<概要>>
::	  TARGET_DIR_PATH 配下（サブフォルダ含む）にあるファイル名
::	  KEY_FILE_NAME を探し、それが格納されたフォルダパスを
::	   システム環境変数の「Path」に追加する
::	
::	<<使い方>>
::	  １．KEY_FILE_NAME に指定した名前のファイルを、システム環境変数の
::		  「Path」に追加したいフォルダ内にコピーする
::	  ２．本バッチファイルを「管理者として実行」する

:: ### 設定情報 ###
set ADD_TO_PATH_SCRIPT_PATH=C:\codes\vbs\700_AddToPathOfEnvVariable.vbs
set TARGET_DIR_PATH=C:\prg_exe\
set KEY_FILE_NAME=_add_to_sys_env_directory

:: ### 処理 ###
FOR /R "%TARGET_DIR_PATH%" %%i IN (%KEY_FILE_NAME%) DO (
	if exist %%i (
		echo %%~dpi
		call %ADD_TO_PATH_SCRIPT_PATH% %%~dpi
	)
)
pause
