@echo off
echo #####################################################################
echo ###  �{�t�H���_�z���ɂ��� csv/txt/htm �ȊO�̃t�@�C����S�č폜���܂��B
set /p EXECUTE_SELECT="###  ���s���܂����H[y/n] : "
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
	echo ###  ���s�𒆎~���܂��I
)
echo #####################################################################
pause
