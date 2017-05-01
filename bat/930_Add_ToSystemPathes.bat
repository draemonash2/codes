@echo off
::	<<概要>>
::	  TARGET_DIR_PATH 配下（サブフォルダ含む）にあるファイル
::	  「_add_to_path.bat」を探し、それが格納されたフォルダパスを
::	   システム環境変数の「Path」に追加する
::	
::	<<使い方>>
::	  １．以下のファイルを作成
::			ファイル名：_add_to_path.bat
::			ファイルの中身：call %1 %~dp0
::	  ２．１で作成したファイルを Path に追加したいフォルダ内にコピーする
::	  ３．本バッチファイルを「管理者として実行」する

:: ### 設定情報 ###
set ADD_TO_PATH_SCRIPT_PATH=C:\codes\vbs\700_AddToPathOfEnvVariable.vbs
set TARGET_DIR_PATH=C:\prg_exe\

:: ### 処理 ###
FOR /R "%TARGET_DIR_PATH%" %%i IN (_add_to_path.bat) DO (
	if exist %%i (
		echo %%i
		call %%i %ADD_TO_PATH_SCRIPT_PATH%
	)
)
pause
