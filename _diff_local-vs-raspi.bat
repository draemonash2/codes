echo リモートのファイルと比較します
	set USER=pi
	set PW=Endo4353
	set LOGINTRGT=raspberrypi.local

echo ファイル取得
	start %MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "get .vimrc %~dp0\" "exit"
	start %MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "get .bashrc %~dp0\" "exit"
	start %MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open %USER%:%PW%@%LOGINTRGT%" "get .inputrc %~dp0\" "exit"
	pause

echo ファイルバックアップ
	copy "%~dp0.vimrc" "%~dp0.vimrc_org"
	copy "%~dp0.bashrc" "%~dp0.bashrc_org"
	copy "%~dp0.inputrc" "%~dp0.inputrc_org"
	pause

echo winmerge比較
	start %MYEXEPATH_WINMERGE% "%~dp0vim\_gvimrc" "%~dp0.vimrc"
	start %MYEXEPATH_WINMERGE% "%~dp0vim\_vimrc" "%~dp0.vimrc"
	start %MYEXEPATH_WINMERGE% "%~dp0vim\.vimrc" "%~dp0.vimrc"
	start %MYEXEPATH_WINMERGE% "%~dp0linux\.bashrc" "%~dp0.bashrc"
	start %MYEXEPATH_WINMERGE% "%~dp0linux\.inputrc" "%~dp0.inputrc"
	pause

echo ファイル送信
	%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open pi:Endo4353@raspberrypi.local" "cd" "put %~dp0.vimrc" "exit"
	%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open pi:Endo4353@raspberrypi.local" "cd" "put %~dp0.bashrc" "exit"
	%MYDIRPATH_PRG_EXE%\WinSCP\WinSCP.exe /console /command "option batch on" "open pi:Endo4353@raspberrypi.local" "cd" "put %~dp0.inputrc" "exit"

echo ファイル削除
	del "%~dp0.vimrc"
	del "%~dp0.bashrc"
	del "%~dp0.inputrc"
	del "%~dp0.vimrc_org"
	del "%~dp0.bashrc_org"
	del "%~dp0.inputrc_org"
	pause
