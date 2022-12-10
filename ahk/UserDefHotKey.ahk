	#NoEnv							; 通常、値が割り当てられていない変数名を参照しようとしたとき、システムの環境変数に同名の変数がないかを調べ、
									; もし存在すればその環境変数の値が参照される。スクリプト中に #NoEnv を記述することにより、この動作を無効化できる。
;	#NoTrayIcon						; スクリプトのタスクトレイアイコンを非表示にする。
;	#Warn							; Enable warnings to assist with detecting common errors.
	SendMode Input					; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。
;	SetWorkingDir %A_ScriptDir%		; スクリプトの作業ディレクトリを本スクリプトの格納ディレクトリに変更。
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え
	#HotkeyModifierTimeout 100		; キーボードフックなしでホットキー中でSendコマンドを使用したときに修飾キーの状態を復元しなくなるタイムアウト時間を設定。
	#WinActivateForce				; ウィンドウのアクティブ化時に、穏やかな方法を試みるのを省略して常に強制的な方法でアクティブ化を行う。（タスクバーアイコンが点滅する現象が起こらなくなる）

	#Include %A_ScriptDir%\lib\IME.ahk

;* ***************************************************************
;* Settings
;* ***************************************************************
DOC_DIR_PATH = C:\Users\%A_Username%\Dropbox\100_Documents
global iWIN_TILE_MODE_CLEAR_INTERVAL := 10000 ; [ms]
global iWIN_TILE_MODE_MAX := 3
global iWIN_Y_OFFSET := 2/7
global iWIN_TILE_MODE_OFFSET := 0

;* ***************************************************************
;* Define variables
;* ***************************************************************
global giWinTileMode := 0

;* ***************************************************************
;* Timer
;* ***************************************************************
	SetTimer ClearWinTileMode, %iWIN_TILE_MODE_CLEAR_INTERVAL%
		Return
	ClearWinTileMode:
		giWinTileMode := iWIN_TILE_MODE_MAX
	;	TrayTip, タイマーClearWinTileMode実行, giWinTileMode = %giWinTileMode%, 1, 17
		Return

;* ***************************************************************
;* Keys
;*  [参考URL]
;*		https://sites.google.com/site/autohotkeyjp/reference/KeyList
;*			無変換）vk1Dsc07B
;*			変換）	vk1Csc079
;*			^）		Control
;*			+）		Shift
;*			!）		Alt
;*			#）		Windowsロゴキー
;* ***************************************************************

;***** キー置き換え *****
	;無変換キー＋方向キーでPgUp,PgDn,Home,End
		vk1Dsc07B::vk1Dsc07B
		vk1Dsc07B & Right::	MuhenkanSimultPush( "End" )
		vk1Dsc07B & Left::	MuhenkanSimultPush( "Home" )
		vk1Dsc07B & Up::	MuhenkanSimultPush( "PgUp" )
		vk1Dsc07B & Down::	MuhenkanSimultPush( "PgDn" )
	;Insertキー
		Insert::Return
	;PrintScreenキー
		PrintScreen::return

;***** ホットキー(Global) *****
	;UserDefHotKey.ahk
		^+!a::
			EnvGet, sExePath, MYEXEPATH_GVIM
			sFilePath = "%A_ScriptFullPath%"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;ホットキー配置表示
		!^+F1::
			sFilePath = "C:\other\グローバルホットキー配置.vsdx"
			StartProgramAndActivateFile( sFilePath )
			return
	;Programsフォルダ表示
		!^+F12::
			sFilePath = "C:\Users\%A_Username%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
			StartProgramAndActivateFile( sFilePath )
			Sleep 100
			Send, +{tab}
			return
	;#todo.itmz
		^+!Up::
			sFilePath = "%DOC_DIR_PATH%\#todo.itmz"
			Process, wait, Dropbox.exe, 30 ; Dropboxが起動(≒同期が完了)するまで待つ(タイムアウト時間30s)
			StartProgramAndActivateFile( sFilePath )
			return
	;#temp.txt
		^+!Down::
			sFilePath = "%DOC_DIR_PATH%\#temp.txt"
			StartProgramAndActivateFile( sFilePath )
			return
	;#temp.xlsm
		^+!Right::
			sFilePath = "%DOC_DIR_PATH%\#temp.xlsm"
			StartProgramAndActivateFile( sFilePath )
			return
	;#temp.vsdm
		^+!Left::
			sFilePath = "%DOC_DIR_PATH%\#temp.vsdm"
			StartProgramAndActivateFile( sFilePath )
			return
	;予算管理.xlsm
		^+!\::
			sFilePath = "%DOC_DIR_PATH%\210_【衣食住】家計\100_予算管理.xlsm"
			StartProgramAndActivateFile( sFilePath )
			return
	;予算管理＠家族用.xlsx
		^+!^::
			sFilePath = "%DOC_DIR_PATH%\..\000_Public\家計\予算管理＠家族用.xlsx"
			StartProgramAndActivateFile( sFilePath )
			return
	;言語チートシート
		^+!c::
			sFilePath = "C:\other\言語チートシート.xlsx"
			StartProgramAndActivateFile( sFilePath )
			return
	;ショートカットキー
		^+!s::
			sFilePath = "C:\other\ショートカットキー一覧.xlsx"
			StartProgramAndActivateFile( sFilePath )
			return
	;$object.xlsm
		^+!o::
			sFilePath = "C:\other\template\#object.xlsm"
			StartProgramAndActivateFile( sFilePath )
			return
	;用語集
		^+!/::
			sFilePath = "%DOC_DIR_PATH%\320_【自己啓発】勉強\words.itmz"
			StartProgramAndActivateFile( sFilePath )
			return
	;KitchenTimer.vbs
		^+!k::
			EnvGet, sDirPath, MYDIRPATH_CODES
			sFilePath = %sDirPath%\vbs\tools\win\other\KitchenTimer.vbs
			StartProgramAndActivateFile( sFilePath )
			return
	;SCPデータ取得
		^+!g::
			EnvGet, sDirPath, MYDIRPATH_CODES
			sFilePath = %sDirPath%\bat\tools\file_ope\FetchScpFromRemote.bat
			StartProgramAndActivateFile( sFilePath )
			return
	;定期キー送信
		^+!t::
			EnvGet, sDirPath, MYDIRPATH_CODES
			sFilePath = %sDirPath%\vbs\tools\win\other\PeriodicKeyTransmission.bat
			StartProgramAndActivateFile( sFilePath )
			return
	;rapture.exe
		^+!x::
			EnvGet, sExePath, MYEXEPATH_RAPTURE
			StartProgramAndActivateExe( sExePath )
			return
	/*
	;xf.exe
		^+!z::
			EnvGet, sExePath, MYEXEPATH_XF
			StartProgramAndActivateExe( sExePath )
			return
	*/
	;DOC_DIR_PATHフォルダ表示
		!^+z::
			sFilePath = "%DOC_DIR_PATH%"
			StartProgramAndActivateFile( sFilePath )
			Sleep 100
			Send, +{tab}
			return
	;cCalc.exe
		^+!;::
			EnvGet, sExePath, MYEXEPATH_CALC
			StartProgramAndActivateExe( sExePath )
			return
	;Github.io
		^+!1::Run https://draemonash2.github.io/
		^+!2::Run https://draemonash2.github.io/linux_sft/linux.html
		^+!3::Run https://draemonash2.github.io/gitcommand_lng/gitcommand.html
	;翻訳サイト
		^+!h::
		;	Run https://translate.google.com/?sl=en&tl=ja&op=translate&hl=ja
			Run https://www.deepl.com//translator
			Return
	;Wifi接続(Bluetoothテザリング起動)
		/*
		^+!w::
			Run, control printers
			Sleep 2000
			Send, myp
			Sleep 300
			Send, {AppsKey}
			Sleep 200
			Send, c
			Sleep 200
			Send, a
			Sleep 5000
			Send, !{F4}
		*/
	;Wifi接続(Wifiテザリング)
		/*
		^+!w::
			EnvGet, sDirPath, MYDIRPATH_CODES
			Run % sDirPath . "\bat\tools\other\ConnectWifi.bat MyPerfectiPhone"
			return
		*/
	;Window最前面化
		Pause::
			;HP製PCでは「Pause」は「Fn＋Shift」。
			WinSet, AlwaysOnTop, TOGGLE, A
			WinGetTitle, sActiveWinTitle, A
			if bEnableAlwaysOnTop = 
			{
				MsgBox, 0x43000, Window最前面化, Window最前面を【有効】にします`n`n%sActiveWinTitle%, 5
				bEnableAlwaysOnTop = 1
			}
			else
			{
				if bEnableAlwaysOnTop = 0
				{
					MsgBox, 0x43000, Window最前面化, Window最前面を【有効】にします`n`n%sActiveWinTitle%, 5
					bEnableAlwaysOnTop = 1
				}
				else
				{
					MsgBox, 0x43000, Window最前面化, Window最前面を【解除】します`n`n%sActiveWinTitle%, 5
					bEnableAlwaysOnTop = 0
				}
			}
			Return
	;Windowタイル切り替え
		!#LEFT::
			IncrementWinTileMode()
			ApplyWinTileMode()
			return
		!#RIGHT::
			DecrementWinTileMode()
			ApplyWinTileMode()
			return
	;Teams一時退席抑止機能
		/*
		+^!F11::
			TrayTip, Teams一時退席抑止機能, Teamsの一時退席を抑止します。`nEscキー長押し(3秒以上)で停止できます。, 5, 17
			Loop
			{
				Sleep, 3000
				GetKeyState, sPressState, Esc, P
				If sPressState = D
				{
					TrayTip, Teams一時退席抑止機能, Teamsの一時退席抑止を解除します。, 5, 17
					Break
				}
				Else
				{
					Send, {vkF3sc029}
				}
			}
			return
		*/
	;テスト用
		/*
		^Pause::
			MsgBox, ctrlpause
			Return
		+Pause::
			MsgBox, shiftpause
			Return
		+^!9::StartProgramAndActivate( "", "C:\Users\draem\Dropbox\100_Documents\#temp.txt" )
		+^!8::StartProgramAndActivate( "C:\Users\draem\Programs\program\prg_exe\Hidemaru\Hidemaru.exe", "C:\Users\draem\Dropbox\100_Documents\#temp.txt" )
		+^!7::StartProgramAndActivate( "C:\Users\draem\Programs\program\prg_exe\Hidemaru\Hidemaru.exe", "" )
		+^!6::StartProgramAndActivate( "", "" )
		+^!9::StartProgramAndActivate( "", "C:\Users\draem\Dropbox\100_Documents\#temp.txt", 0 )
		+^!8::StartProgramAndActivate( "C:\Users\draem\Programs\program\prg_exe\Hidemaru\Hidemaru.exe", "C:\Users\draem\Dropbox\100_Documents\#temp.txt", 0 )
		+^!7::StartProgramAndActivate( "C:\Users\draem\Programs\program\prg_exe\Hidemaru\Hidemaru.exe", "", 0 )
		+^!6::StartProgramAndActivate( "", "", 0 )
		+^!9::StartProgramAndActivate( "", "C:\Users\draem\Dropbox\100_Documents\#temp.txt", 1 )
		+^!8::StartProgramAndActivate( "C:\Users\draem\Programs\program\prg_exe\Hidemaru\Hidemaru.exe", "C:\Users\draem\Dropbox\100_Documents\#temp.txt", 1 )
		+^!7::StartProgramAndActivate( "C:\Users\draem\Programs\program\prg_exe\Hidemaru\Hidemaru.exe", "", 1 )
		+^!6::StartProgramAndActivate( "", "", 1 )
		^1::
			MouseGetPos,x,y,hwnd,ctrl,3
			MouseClick, left, 1209, 932
			Sleep 100
			MouseClick, left, 1127, 1184
			Sleep 100
			MouseClick, left, 2089, 302
			Sleep 100
			MouseMove, x, y
			return
		*/

;***** ホットキー(Software local) *****
	#IfWinNotActive ahk_exe WindowsTerminal.exe
		RAlt::Send, {AppsKey}	;右Altキーをコンテキストメニュー表示に変更
	#IfWinNotActive
	
	#IfWinActive ahk_exe explorer.exe
		^+c::	; ファイルパスコピー
			sTrgtPaths := GetSelFilePathAtExplorer(0)
			Clipboard =
			Clipboard = %sTrgtPaths%
			ClipWait
			return
		^+d::	; ファイル名コピー
			sTrgtNames := GetSelFileNameAtExplorer()
			Clipboard =
			Clipboard = %sTrgtNames%
			ClipWait
			return
		+F1::	; winmergeで開く
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			EnvGet, sDirPath, MYDIRPATH_CODES
			Run, % sDirPath . "\vbs\tools\wimmerge\CompareWithWinmerge.vbs " . sTrgtPaths
			return
		+F2::	; vimで開く
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			EnvGet, sExePath, MYEXEPATH_GVIM
			StartProgramAndActivate( sExePath, sTrgtPaths )
			return
		+F3::	; VSCodeで開く
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			EnvGet, sExePath, MYEXEPATH_VSCODE
			StartProgramAndActivate( sExePath, sTrgtPaths )
			return
		+F4::	; 秀丸で開く
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			EnvGet, sExePath, MYEXEPATH_HIDEMARU
			StartProgramAndActivate( sExePath, sTrgtPaths )
			return
		+F5::	; EXCELで開く
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			EnvGet, sExePath, MYEXEPATH_EXCEL
			StartProgramAndActivate( sExePath, sTrgtPaths )
			return
		^+g::	; Grep検索＠TresGrep
			sTrgtPaths := GetCurDirPathAtExplorer()
			EnvGet, sExePath, MYEXEPATH_TRESGREP
			StartProgramAndActivate( sExePath, sTrgtPaths )
			return
		^+z::	; 圧縮/パスワード圧縮/解凍
			Gui, New, ,
			Gui, Add, Text,, 圧縮/パスワード圧縮/解凍を実行します。`n処理を選択してください。
			Gui, Add, ListBox, vAnswer Choose1 R3, 圧縮|パスワード圧縮|解凍
			Gui, Add, Button, Hidden w0 h0 Default, ZipEnter
			Gui, Show, Center
			return
			
			ButtonZipEnter:
				GuiControlGet, vAnswer, , ListBox1
				;MsgBox %vAnswer%
				Gui, Cancel
				sTrgtPaths := GetSelFilePathAtExplorer(1)
				EnvGet, sDirPath, MYDIRPATH_CODES
				If ( vAnswer == "圧縮" ) {
					Run, % sDirPath . "\vbs\tools\7zip\ZipFile.vbs " . sTrgtPaths
				} Else If ( vAnswer == "パスワード圧縮" ) {
					Run, % sDirPath . "\vbs\tools\7zip\ZipPasswordFile.vbs " . sTrgtPaths
				} Else If ( vAnswer == "解凍" ) {
					Run, % sDirPath . "\vbs\tools\7zip\UnzipFile.vbs " . sTrgtPaths
				} Else {
					MsgBox "[ERROR] 圧縮/パスワード圧縮/解凍 選択"
				}
				return
		^+l::	; ショートカット/シンボリックリンク作成
			Gui, New, ,
			Gui, Add, Text,, ショートカット/シンボリックリンクを作成します。`n処理を選択してください。
			Gui, Add, ListBox, vAnswer Choose1 R2, ショートカット作成|シンボリックリンク作成
			Gui, Add, Button, Hidden w0 h0 Default, SelLinkEnter
			Gui, Show, Center
			return
			
			ButtonSelLinkEnter:
				GuiControlGet, vAnswer, , ListBox1
				;MsgBox %vAnswer%
				Gui, Cancel
				sTrgtPaths := GetSelFilePathAtExplorer(1)
				EnvGet, sDirPath, MYDIRPATH_CODES
				If ( vAnswer == "ショートカット作成" ) {
					Run % sDirPath . "\vbs\command\CreateShortcutFile.vbs " . sTrgtPaths . ".lnk " . sTrgtPaths
				} Else If ( vAnswer == "シンボリックリンク作成" ) {
					Run % sDirPath . "\vbs\tools\win\file_ope\CreateSymbolicLink.vbs " . sTrgtPaths
				} Else {
					MsgBox "[ERROR] ショートカット/シンボリックリンク作成"
				}
				return
		^+r::	; リネーム用バッチファイル作成
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			EnvGet, sDirPath, MYDIRPATH_CODES
			Run, % sDirPath . "\vbs\tools\win\file_ope\CreateRenameBat.vbs " . sTrgtPaths
			return
		^+F3::	; 隠しファイル 表示非表示切替え
			Send, !vhh
			return
		^+F10::	; コマンドプロンプトを開く
			sDirPath := GetCurDirPathAtExplorer()
			Run, %comspec% /k cd %sDirPath%
			return
		^+F11::	; パス一覧作成
			Gui, New, ,
			Gui, Add, Text,, パス一覧を作成します。`n処理を選択してください。
			Gui, Add, ListBox, vAnswer Choose1 R4, ファイル＆フォルダ一覧作成|ファイル一覧作成|フォルダ一覧作成|フォルダツリー作成
			Gui, Add, Button, Hidden w0 h0 Default, PathListEnter
			Gui, Show, Center
			return
			
			ButtonPathListEnter:
				GuiControlGet, vAnswer, , ListBox1
				;MsgBox %vAnswer%
				Gui, Cancel
				sDirPath := GetCurDirPathAtExplorer()
				If ( vAnswer == "ファイル＆フォルダ一覧作成" ) {
					Run, %ComSpec% /c dir /s /b /a > "%sDirPath%\_PathList_FileDir.txt", %sDirPath%
				} Else If ( vAnswer == "ファイル一覧作成" ) {
					Run, %ComSpec% /c dir *.* /b /s /a:a-d > "%sDirPath%\_PathList_File.txt", %sDirPath%
				} Else If ( vAnswer == "フォルダ一覧作成" ) {
					Run, %ComSpec% /c dir /b /s /a:d > "%sDirPath%\_PathList_Dir.txt", %sDirPath%
				} Else If ( vAnswer == "フォルダツリー作成" ) {
					Run, %ComSpec% /c tree /f > "%sDirPath%\_DirTree.txt", %sDirPath%
				} Else {
					MsgBox "[ERROR] パス一覧作成"
				}
				return
		^+F12::	; フォルダサイズ解析＠DiskInfo
			sTrgtPaths := GetCurDirPathAtExplorer()
			EnvGet, sExePath, MYEXEPATH_DISKINFO3
			StartProgramAndActivate( sExePath, sTrgtPaths )
			return
		
		GuiEscape:
		GuiClose:
		ButtonCancel:
			Gui, Cancel
			Return
	#IfWinActive
	
	#IfWinActive ahk_exe EXCEL.EXE
		F1::return	;F1ヘルプ無効化
		+Space::	;IME ON状態でShift+Space(行選択)が効かない対策
			if (IME_GET() == 1) {
				IME_SET(0)
				Sleep 50
				SendInput +{Space}
				Sleep 50
				IME_SET(1)
			} else {
				SendInput +{Space}
			}
			Return
	#IfWinActive
	
	#IfWinActive ahk_exe iThoughts.exe
		F1::return	;F1ヘルプ無効化
	#IfWinActive
	
	#IfWinActive ahk_exe Rapture.exe
		Esc::!F4	;Escで終了
	#IfWinActive
	
	#IfWinActive ahk_exe vimrun.exe
		Esc::!F4	;Escで終了
	#IfWinActive
	
	#IfWinActive ahk_exe XF.exe
		^WheelUp::SendInput ^+{Tab}  ;Next tab.
		^WheelDown::SendInput ^{Tab} ;Previous tab.
	#IfWinActive
	
	#IfWinActive ahk_exe chrome.exe
	;	^WheelUp::SendInput ^+{Tab}  ;Next tab.
	;	^WheelDown::SendInput ^{Tab} ;Previous tab.
	#IfWinActive
	
	#IfWinActive ahk_class MPC-BE
		]::Send, {Space}
	#IfWinActive
	
	;#IfWinActive ahk_exe Kindle.exe
		;Kindle 自動ページ送り
		/*
		bIsAutoPageFeed=0
		^+!9::
			If (bIsAutoPageFeed=0)
			{
				MsgBox 自動ページ送りを起動します
				bIsAutoPageFeed=1
				SetTimer, AutoPageFeed, 3000
			}
			Else
			{
				MsgBox 自動ページ送りを無効化します
				bIsAutoPageFeed=0
				SetTimer, AutoPageFeed, Off
			}
			Return
		AutoPageFeed:
			IfWinActive ahk_exe Kindle.exe
			{
				Send, {Right}
			}
			Return
		*/
	;#IfWinActive
	
	#IfWinActive ahk_exe PDFXEdit.exe
		MButton::	SendInput ^z ;元に戻す
		XButton1::	SendInput !5 ;下線
		XButton2::	SendInput !4 ;テキストハイライト
		;ハイライトを既定の書式設定に変更する
		+^!F11::
			Loop, 20
			{
				Sleep 200
				Send, !{Enter}
				Sleep 200
				Send, !a
				Sleep 200
				Send, {Up}
				Sleep 200
				Send, {Enter}
				Sleep 300
				Send, {Tab}
				Sleep 200
				Send, {Enter}
				Sleep 300
				Send, {Down}
			}
			MsgBox 完了！
			Return
	#IfWinActive

;* ***************************************************************
;* Functions
;* ***************************************************************
	; 起動＆アクティベート処理
	;
	; 既定のショートカットキーとの干渉によりプログラム起動後に
	; ウィンドウがアクティベートされないことがある。(※)
	; 上記問題を対処するため、本関数ではプログラム起動後に
	; ウィンドウをアクティベートする処理を実行する。
	;
	; (※)例
	; 「Windows キー + 1」はタスクパーに１つ目にピン止め
	; されているプログラムをアクティベートするショートカットキーで
	; あるため、Run 関数を使用してそのまま実行すると、非アクティブ
	; 状態でプログラムが起動してしまう。
	StartProgramAndActivate( sExePath, sFilePath )
	{
		;*** preprocess ***
		If ( sExePath == "" or sFilePath == "" )
		{
			MsgBox [ERROR] please specify arguments to StartProgramAndActivate().
			return
		}
		sExeDirPath := ExtractDirPath(sExePath)
		;MsgBox sExePath=%sExePath% `n sExeDirPath=%sExeDirPath% `n sFilePath=%sFilePath%
		
		;*** start program ***
		Try {
			Run, %sExePath% %sFilePath%, %sExeDirPath%, , sOutputVarPID
		;	MsgBox, 0x1000, , %sOutputVarPID%
		} Catch errorno {
			MsgBox, [error] run error : %errorno%
		}
		WinActivate, ahk_pid %sOutputVarPID%
		return
	}
	
	; 起動＆アクティベート処理 (ファイルパス指定のみ)
	StartProgramAndActivateFile( sFilePath )
	{
		;*** preprocess ***
		If ( sFilePath == "" )
		{
			MsgBox [ERROR] please specify arguments to StartProgramAndActivateFile().
			return
		}
		sFileName := ExtractFileName(sFilePath)
		;MsgBox sFilePath=%sFilePath% `n sFileName=%sFileName%
		
		;*** start program ***
		Try {
			Run, %sFilePath%, , , sOutputVarPID
		;	MsgBox, 0x1000, , %sOutputVarPID%
		} Catch errorno {
			MsgBox, [error] run error : %errorno%
		}
		WinActivate, ahk_pid %sOutputVarPID%
		return
	}
	
	; 起動＆アクティベート処理 (実行プログラム指定のみ)
	;   "sExePathのみ指定"かつ"起動済み"の場合はアクティブ化のみを行う
	StartProgramAndActivateExe( sExePath )
	{
		;*** preprocess ***
		If ( sExePath == "" )
		{
			MsgBox [ERROR] please specify arguments to StartProgramAndActivateExe().
			return
		}
		
		sExeName := ExtractFileName(sExePath)
		sExeDirPath := ExtractDirPath(sExePath)
		;MsgBox sExePath=%sExePath% `n sExeDirPath=%sExeDirPath% `n sExeName=%sExeName%
		
		;*** start program ***
		Process, Exist, % sExeName
		If ErrorLevel<>0
		{
			WinActivate,ahk_pid %ErrorLevel%
		}
		Else
		{
			Try {
				Run, %sExePath%, %sExeDirPath%, , sOutputVarPID
			;	MsgBox, 0x1000, , %sOutputVarPID%
			} Catch errorno {
				MsgBox, [error] run error : %errorno%
			}
			WinActivate, ahk_pid %sOutputVarPID%
		}
		return
	}

	; 無変換キー同時押し実装
	MuhenkanSimultPush( sSendKey )
	{
		if(GetKeyState("Shift","P") and GetKeyState("Ctrl","P") and GetKeyState("Alt","P")){
			Send !^+{%sSendKey%}
		} else if(GetKeyState("Shift","P") and GetKeyState("Ctrl","P")){
			Send ^+{%sSendKey%}
		} else if(GetKeyState("Shift","P") and GetKeyState("Alt","P")){
			Send !+{%sSendKey%}
		} else if(GetKeyState("Alt","P") and GetKeyState("Ctrl","P")){
			Send !^{%sSendKey%}
		} else if(GetKeyState("Alt","P")){
			Send !{%sSendKey%}
		} else if(GetKeyState("Ctrl","P")){
			Send ^{%sSendKey%}
		} else if(GetKeyState("Shift","P")){
			Send +{%sSendKey%}
		} else {
			Send {%sSendKey%}
		}
		return
	}

	;Windowタイル切り替え
	GetWinTileModeMin()
	{
		SysGet, iMonitorNum, MonitorCount
		if (iMonitorNum = 2) {
			iWinTileModeMin := 0
		} else {
			iWinTileModeMin := 3
		}
	;	MsgBox, iMonitorNum: %iMonitorNum%`ngiWinTileModeMin: %iWinTileModeMin%
		return iWinTileModeMin
	}
	IncrementWinTileMode()
	{
		iWinTileModeMin := GetWinTileModeMin()
		if ( giWinTileMode >= iWIN_TILE_MODE_MAX ) {
			giWinTileMode := iWinTileModeMin
		} else {
			giWinTileMode++
		}
	;	MsgBox, giWinTileMode: %giWinTileMode%`n iWIN_TILE_MODE_MAX: %iWIN_TILE_MODE_MAX%`n iWinTileModeMin: %iWinTileModeMin%
	}
	DecrementWinTileMode()
	{
		iWinTileModeMin := GetWinTileModeMin()
		if ( giWinTileMode <= iWinTileModeMin ) {
			giWinTileMode := iWIN_TILE_MODE_MAX
		} else {
			giWinTileMode--
		}
	;	MsgBox, giWinTileMode: %giWinTileMode%`n iWIN_TILE_MODE_MAX: %iWIN_TILE_MODE_MAX%`n iWinTileModeMin: %iWinTileModeMin%
	}
	GetMonitorPosInfo( MonitorNum, ByRef X, ByRef Y, ByRef Width, ByRef Height )
	{
		SysGet, Mon, MonitorWorkArea, %MonitorNum%
	;	MsgBox, Left:%MonLeft%`nRight:%MonRight%`nTop:%MonTop%`nBottom:%MonBottom%
	;	SysGet, MonName, MonitorName, %MonitorNum%
	;	MsgBox, MonName:%MonName%
		Y:=MonTop
		if ( MonLeft < MonRight ) {
			X:=MonLeft
			Width:= % MonRight - MonLeft + 1
		} else {
			X:=% MonRight
			Width:= % MonLeft - MonRight + 1
		}
		Height:= % MonBottom - MonTop + 1
	;	MsgBox, %MonitorNum%`n%X%`n%Y%`n%Width%`n%Height%
	}
	; ウィンドウサイズ切り替え
	ApplyWinTileMode()
	{
		GetMonitorPosInfo(1, mainx, mainy, mainwidth, mainheight )
		GetMonitorPosInfo(2, subx, suby, subwidth, subheight )
	;	MsgBox, mainx: %mainx%`nmainy: %mainy%`nmainwidth: %mainwidth%`nmainheight: %mainheight%`nsubx: %subx%`nsuby: %suby%`nsubwidth: %subwidth%`nsubheight: %subheight%
		
		winywhole:= % suby + ( subheight * iWIN_Y_OFFSET )
		winheightwhole:= % subheight * ( 1 - iWIN_Y_OFFSET )
	;	MsgBox, giWinTileMode: %giWinTileMode%`nwinywhole: %winywhole%`nwinheightwhole: %winheightwhole%
		if ( giWinTileMode = 0 ) {			;サブ全体
			winx:=subx
			winwidth:=subwidth
			winy:=winywhole
			winheight:=winheightwhole
		} else if ( giWinTileMode = 1 ) {	;サブ上
			winx:=subx
			winwidth:=subwidth
			winy:=winywhole
			winheight:= % winheightwhole // 2
		} else if ( giWinTileMode = 2 ) {	;サブ下
			winx:=subx
			winwidth:=subwidth
			winy:= % winywhole + winheightwhole // 2
			winheight:=% winheightwhole // 2
		} else if ( giWinTileMode = 3 ) {	;メイン全体
			winx:=mainx
			winy:=mainy
			winwidth:=mainwidth
			winheight:=mainheight
		} else if ( giWinTileMode = 4 ) {	;メイン左
			winx:=% mainx - iWIN_TILE_MODE_OFFSET
			winy:=mainy
			winwidth:=% mainwidth // 2 + iWIN_TILE_MODE_OFFSET
			winheight:=% mainheight + iWIN_TILE_MODE_OFFSET
		} else if ( giWinTileMode = 5 ) {	;メイン右
			winx:=% mainx + mainwidth // 2 - iWIN_TILE_MODE_OFFSET
			winy:=mainy
			winwidth:=% mainwidth // 2 + iWIN_TILE_MODE_OFFSET
			winheight:=% mainheight + iWIN_TILE_MODE_OFFSET
		} else {
			MsgBox, [error] invalid giWinTileMode.`n %giWinTileMode%
			return
		}
	;	MsgBox, giWinTileMode: %giWinTileMode%`nwinx: %winx%`nwiny: %winy%`nwinwidth: %winwidth%`nwinheight: %winheight%
		WinMove, A, , winx, winy, winwidth, winheight
		return
	}
	
	; ファイル名取得
	ExtractFileName( sFilePath )
	{
		SplitPath, sFilePath, sFileName, sDirPath, sExtName, sFileBaseName, sDrive
		StringReplace, sFileName, sFileName, ", , All
	;	MsgBox, %sFilePath%`n%sFileName%`n%sDirPath%`n%sExtName%`n%sFileBaseName%`n%sDrive%
		return sFileName
	}
	; ディレクトリパス取得
	ExtractDirPath( sFilePath )
	{
		SplitPath, sFilePath, sFileName, sDirPath, sExtName, sFileBaseName, sDrive
		StringReplace, sDirPath, sDirPath, ", , All
	;	MsgBox, %sFilePath%`n%sFileName%`n%sDirPath%`n%sExtName%`n%sFileBaseName%`n%sDrive%
		return sDirPath
	}

	; 選択ファイルパス取得＠explorer
	GetSelFilePathAtExplorer( bIsDelimiterSpace )
	{
		clipboard_old=%Clipboard%
		Clipboard =
		Send, !hcp
		ClipWait
		sTrgtPaths = %Clipboard%
		Clipboard = %clipboard_old%
		if ( bIsDelimiterSpace = 1 ) {
			StringReplace, sTrgtPaths, sTrgtPaths, `r`n, %A_Space%, All
		}
	;	MsgBox sTrgtPaths=%sTrgtPaths%
		return sTrgtPaths
	}
	; 現在フォルダパス取得＠explorer
	GetCurDirPathAtExplorer()
	{
		clipboard_old=%Clipboard%
		Clipboard =
		Send, !d
		Send, ^c
		Send, {ESC}
		ClipWait
		sTrgtPaths = %Clipboard%
		Clipboard = %clipboard_old%
	;	MsgBox sTrgtPaths=%sTrgtPaths%
		return sTrgtPaths
	}
	; 選択ファイル名取得＠explorer
	GetSelFileNameAtExplorer()
	{
		sFilePaths := GetSelFilePathAtExplorer(0)
		sDirPaths := GetCurDirPathAtExplorer()
		StringReplace, sTrgtPaths, sFilePaths, %sDirPaths%\, , All
		StringReplace, sTrgtPaths, sTrgtPaths, ", , All
	;	MsgBox sTrgtPaths=%sTrgtPaths% `n sFilePaths=%sFilePaths% `n sDirPaths=%sDirPaths%
		return sTrgtPaths
	}

