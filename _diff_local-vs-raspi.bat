set USER=pi
set PW=Endo4353
set LOGINTRGT=raspberrypi.local
set HOMEDIR=/home/pi

set answer=y
if %answer%==y (
	set /p answer=ファイルを取得します。処理を継続しますか？[y/n]:
	if %answer%==y (
		%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "get %HOMEDIR%/.vimrc %~dp0" "exit"
		%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "get %HOMEDIR%/.bashrc %~dp0" "exit"
		%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "get %HOMEDIR%/.inputrc %~dp0" "exit"
		pause
	)
)

if %answer%==y (
	echo ファイルバックアップ
	copy "%~dp0.vimrc" "%~dp0.vimrc_rmtorg"
	copy "%~dp0.bashrc" "%~dp0.bashrc_rmtorg"
	copy "%~dp0.inputrc" "%~dp0.inputrc_rmtorg"
	pause
)

if %answer%==y (
	echo winmerge比較
	start %MYEXEPATH_WINMERGE% "%~dp0linux\.inputrc" "%~dp0.inputrc"
	start %MYEXEPATH_WINMERGE% "%~dp0linux\.bashrc" "%~dp0.bashrc"
	start %MYEXEPATH_WINMERGE% "%~dp0vim\.vimrc" "%~dp0.vimrc"
	start %MYEXEPATH_WINMERGE% "%~dp0vim\_vimrc" "%~dp0.vimrc"
	start %MYEXEPATH_WINMERGE% "%~dp0vim\_gvimrc" "%~dp0.vimrc"
	pause
)

if %answer%==y (
	set /p answer=ファイルを送信します。処理を継続しますか？[y/n]:
	if %answer%==y (
		%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "cd" "put %~dp0.vimrc" "exit"
		%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "cd" "put %~dp0.bashrc" "exit"
		%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "cd" "put %~dp0.inputrc" "exit"
		pause
	)
)

::if %answer%==y (
::	set /p answer=バックアップファイルを削除します。処理を継続しますか？[y/n]:
::	if %answer%==y (
::		echo ファイル削除
::		del "%~dp0.vimrc"
::		del "%~dp0.bashrc"
::		del "%~dp0.inputrc"
::		del "%~dp0.vimrc_rmtorg"
::		del "%~dp0.bashrc_rmtorg"
::		del "%~dp0.inputrc_rmtorg"
::		pause
::	)
::)

