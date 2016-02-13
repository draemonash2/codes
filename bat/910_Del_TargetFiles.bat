@echo off
echo #####################################################################
echo ###  本フォルダ配下にある csv/txt/htm 以外のファイルを全て削除します。
set /p EXECUTE_SELECT="###  実行しますか？[y/n] : "
echo ###   
if %EXECUTE_SELECT% == y (
	for /r %%i in (*.*) do (
		if %~f0 == %%i (
			echo ###   [no change] %%i
		) else if %%~xi == .csv (
			echo ###   [no change] %%i
		) else if %%~xi == .txt (
			echo ###   [no change] %%i
		) else if %%~xi == .htm (
			echo ###   [no change] %%i
		) else (
			echo ###   [delete   ] %%i
			del %%i
		)
	)
) else (
	echo ###  実行を中止します！
)
echo #####################################################################
pause
