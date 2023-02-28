;	https://www.autohotkey.com/docs/v2/index.htm

;	#NoTrayIcon						; スクリプトのタスクトレイアイコンを非表示にする。
	#Warn All						; Enable warnings to assist with detecting common errors.
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え
	#WinActivateForce				; ウィンドウのアクティブ化時に、穏やかな方法を試みるのを省略して常に強制的な方法でアクティブ化を行う。（タスクバーアイコンが点滅する現象が起こらなくなる）
	SendMode "Input"				; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。


;* ***************************************************************
;* Settings
;* ***************************************************************
global gsDOC_DIR_PATH := "C:\Users\" . A_Username . "\Dropbox\100_Documents"
global iWIN_TILE_MODE_CLEAR_INTERVAL := 10000 ; [ms]
global iWIN_TILE_MODE_MAX := 3
global iWIN_Y_OFFSET := 2/7
global iWIN_TILE_MODE_OFFSET := 0
global bEnableAlwaysOnTop := 0

;* ***************************************************************
;* Define variables
;* ***************************************************************
global giWinTileMode := 0
global Dim := 0
global DimId := 0

;* ***************************************************************
;* Preprocess
;* ***************************************************************
SetTimerWinTileMode()
TraySetIcon "UserDefHotKey2.ico"
DimMon_GenFilter()

;* ***************************************************************
;* Keys
;*  [参考URL]
;*		https://www.autohotkey.com/docs/v2/KeyList.htm
;*			^）		Control
;*			+）		Shift
;*			!）		Alt
;*			#）		Windowsロゴキー
;* ***************************************************************

;***** キー置き換え *****
	;無変換キー＋方向キーでPgUp,PgDn,Home,End
		VK1D::VK1D
		VK1D & Right::	MuhenkanSimultPush( "End" )
		VK1D & Left::	MuhenkanSimultPush( "Home" )
		VK1D & Up::		MuhenkanSimultPush( "PgUp" )
		VK1D & Down::	MuhenkanSimultPush( "PgDn" )
	;Insertキー
		Insert::Return
	;PrintScreenキー
		PrintScreen::return

;***** ホットキー(Global) *****
	;スクリプトリロード
		^+!F5::
		{
			Reload
			Sleep 1000 ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
			Result := MsgBox("The script could not be reloaded. Would you like to open it for editing?",, 4)
			if Result = "Yes"
				Edit
		}
	;UserDefHotKey.ahk
		^+!a::
		{
			sExePath := EnvGet("MYEXEPATH_GVIM")
			sFilePath := A_ScriptFullPath
			StartProgramAndActivate( sExePath, sFilePath )
		}
	;ホットキー配置表示
		!^+F1::
		{
			sFilePath := "C:\other\グローバルホットキー配置.vsdx"
			StartProgramAndActivateFile( sFilePath )
		}
	;Programsフォルダ表示
		!^+F12::
		{
			sFilePath := "C:\Users\" . A_Username . "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
			StartProgramAndActivateFile( sFilePath )
			Sleep 100
			Send "+{tab}"
		}
	;#todo.itmz
		^+!Up::
		{
			sFilePath := gsDOC_DIR_PATH . "\#todo.itmz"
		;	lPID := ProcessWait("Dropbox.exe", 30) ; Dropboxが起動(≒同期が完了)するまで待つ(タイムアウト時間30s)
			StartProgramAndActivateFile( sFilePath )
		}
	;#temp.txt
		^+!Down::
		^+!Space::
		{
			sFilePath := "C:\Users\draem\Dropbox\100_Documents\#temp.txt"
			StartProgramAndActivateFile( sFilePath )
		}
	;#temp.xlsm
		^+!Right::
		{
			sFilePath := gsDOC_DIR_PATH . "\#temp.xlsm"
			StartProgramAndActivateFile( sFilePath )
		}
	;#temp.vsdm
		^+!Left::
		{
			sFilePath := gsDOC_DIR_PATH . "\#temp.vsdm"
			StartProgramAndActivateFile( sFilePath )
		}
	;予算管理.xlsm
		^+!\::
		{
			sFilePath := gsDOC_DIR_PATH . "\210_【衣食住】家計\100_予算管理.xlsm"
			StartProgramAndActivateFile( sFilePath )
		}
	;予算管理＠家族用.xlsx
		^+!^::
		{
			sFilePath := gsDOC_DIR_PATH . "\..\000_Public\家計\予算管理＠家族用.xlsx"
			StartProgramAndActivateFile( sFilePath )
		}
	;言語チートシート
		^+!c::
		{
			sFilePath := "C:\other\言語チートシート.xlsx"
			StartProgramAndActivateFile( sFilePath )
		}
	;ショートカットキー
		^+!s::
		{
			sFilePath := "C:\other\ショートカットキー一覧.xlsx"
			StartProgramAndActivateFile( sFilePath )
		}
	;#object.xlsm
		^+!o::
		{
			sFilePath := "C:\other\template\#object.xlsm"
			StartProgramAndActivateFile( sFilePath )
		}
	;用語集
		^+!/::
		{
			sFilePath := gsDOC_DIR_PATH . "\320_【自己啓発】勉強\words.itmz"
			StartProgramAndActivateFile( sFilePath )
		}
	;codes同期
		^+!y::
		{
			sDirPath := EnvGet("MYDIRPATH_CODES")
			sFilePath := sDirPath . "\_sync_github-codes-remote.bat"
			StartProgramAndActivateFile( sFilePath )
		}
	;KitchenTimer.vbs
		^+!k::
		{
			sDirPath := EnvGet("MYDIRPATH_CODES")
			sFilePath := sDirPath . "\vbs\tools\win\other\KitchenTimer.vbs"
			StartProgramAndActivateFile( sFilePath )
		}
	;定期キー送信
		^+!t::
		{
			sDirPath := EnvGet("MYDIRPATH_CODES")
			sFilePath := sDirPath . "\vbs\tools\win\other\PeriodicKeyTransmission.bat"
			StartProgramAndActivateFile( sFilePath )
		}
	;rapture.exe
		^+!x::
		{
			static DimOld := 0
			global Dim
			; 明るさを最大にする
			DimOld := Dim
			Dim := 0
			DimMon_LoopMonitor()
			; Rapture 起動
			sExePath := EnvGet("MYEXEPATH_RAPTURE")
			StartProgramAndActivateExe( sExePath )
			; 明るさを元に戻す
			Sleep 5000
			Dim := DimOld
			DimMon_LoopMonitor()
		}
	;xf.exe
	/*
		^+!z::
		{
			sExePath := EnvGet("MYEXEPATH_XF")
			StartProgramAndActivateExe( sExePath, 1 )
		}
	*/
	;gsDOC_DIR_PATHフォルダ表示
		!^+z::
		{
			sFilePath := gsDOC_DIR_PATH
			StartProgramAndActivateFile( sFilePath )
			Sleep 100
			Send "+{tab}"
		}
	;cCalc.exe
		^+!;::
		{
			sExePath := EnvGet("MYEXEPATH_CALC")
			StartProgramAndActivateExe( sExePath, 1 )
		}
	;Github.io
		^+!1::Run "https://draemonash2.github.io/"
		^+!2::Run "https://draemonash2.github.io/linux_sft/linux.html"
		^+!3::Run "https://draemonash2.github.io/gitcommand_lng/gitcommand.html"
	;翻訳サイト
		^+!h::
		{
		;	Run "https://translate.google.com/?sl=en&tl=ja&op=translate&hl=ja"
			Run "https://www.deepl.com//translator"
		}
	;Wifi接続(Bluetoothテザリング起動)
		/*
		^+!w::
		{
			Run, control printers
			
			Sleep 2000
			Send "myp"
			Sleep 300
			Send "{AppsKey}"
			Sleep 200
			Send "c"
			Sleep 200
			Send "a"
			Sleep 5000
			Send "!{F4}"
		*/
	;Wifi接続(Wifiテザリング)
		/*
		^+!w::
		{
			sDirPath := EnvGet("MYDIRPATH_CODES")
			Run % sDirPath . "\bat\tools\other\ConnectWifi.bat MyPerfectiPhone"
		}
		*/
	;Window最前面化
		Pause::
		{
			;HP製PCでは「Pause」は「Fn＋Shift」。
			WinSetAlwaysOnTop -1, "A"
			sActiveWinTitle := WinGetTitle("A")
			if (bEnableAlwaysOnTop = 0)
			{
				MsgBox "Window最前面を【有効】にします`n`n" . sActiveWinTitle, "Window最前面化", 0x43000
				global bEnableAlwaysOnTop := 1
			}
			else
			{
				MsgBox "Window最前面を【解除】します`n`n" . sActiveWinTitle, "Window最前面化", 0x43000
				global bEnableAlwaysOnTop := 0
			}
		}
	;Windowタイル切り替え
		!#LEFT::
		{
			SetTimerWinTileMode()
			IncrementWinTileMode()
			ApplyWinTileMode()
		}
		!#RIGHT::
		{
			SetTimerWinTileMode()
			DecrementWinTileMode()
			ApplyWinTileMode()
		}
	;Teams一時退席抑止機能
		/*
		+^!F11::
		{
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
					Send "{vkF3sc029}"
				}
			}
		}
		*/
	;DimMonitor
		#Home::							; 明度100%（不透明度0%）
		{
			global Dim
			Dim := 0
			DimMon_HotKey()
		}
		#End::							; 明度0%（不透明度100%）
		{
			global Dim
			Dim := 80
			DimMon_HotKey()
		}
		#PgDn::							; 明度を下げる（不透明度を上げる）
		{
			global Dim
			Dim += 20
			if (Dim > 80)
				Dim := 80
			DimMon_HotKey()
		}
		
		#PgUp::							; 明度を上げる（不透明度を下げる）
		{
			global Dim
			Dim -= 20
			if (Dim < 0)
				Dim := 0
			DimMon_HotKey()
		}
	;テスト用
		/*
		^Pause::
		{
			MsgBox, ctrlpause
		}
		+Pause::
		{
			MsgBox, shiftpause
		}
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
		{
			MouseGetPos,x,y,hwnd,ctrl,3
			MouseClick, left, 1209, 932
			Sleep 100
			MouseClick, left, 1127, 1184
			Sleep 100
			MouseClick, left, 2089, 302
			Sleep 100
			MouseMove, x, y
		}
		*/
		
		^1::
		{
			KeyHistory
		}

;***** ホットキー(Software local) *****
	#HotIf !WinActive("ahk_exe WindowsTerminal.exe")
		RAlt::Send "{AppsKey}"	;右Altキーをコンテキストメニュー表示に変更
	#HotIf
	
	#HotIf WinActive("ahk_exe explorer.exe")
		^+c::	; ファイルパスコピー
		{
			sTrgtPaths := GetSelFilePathAtExplorer(0)
			A_Clipboard := ""
			A_Clipboard := sTrgtPaths
			ClipWait
			;FocusFileDirListAtExplorer()
		}
		^+d::	; ファイル名コピー
		{
			sTrgtNames := GetSelFileNameAtExplorer()
			A_Clipboard := ""
			A_Clipboard := sTrgtNames
			ClipWait
			;FocusFileDirListAtExplorer()
		}
		+F1::	; winmergeで開く
		{
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			sDirPath := EnvGet("MYDIRPATH_CODES")
			Run sDirPath . "\vbs\tools\wimmerge\CompareWithWinmerge.vbs " . sTrgtPaths
		}
		+F2::	; vimで開く
		{
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			sExePath := EnvGet("MYEXEPATH_GVIM")
			StartProgramAndActivate( sExePath, sTrgtPaths )
		}
		+F3::	; VSCodeで開く
		{
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			sExePath := EnvGet("MYEXEPATH_VSCODE")
			StartProgramAndActivate( sExePath, sTrgtPaths )
		}
		+F4::	; 秀丸で開く
		{
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			sExePath := EnvGet("MYEXEPATH_HIDEMARU")
			StartProgramAndActivate( sExePath, sTrgtPaths )
		}
		+F5::	; EXCELで開く
		{
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			sExePath := EnvGet("MYEXEPATH_EXCEL")
			StartProgramAndActivate( sExePath, sTrgtPaths )
		}
		+F9::	; 作業ファイルとしてコピー
		{
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			sDirPath := EnvGet("MYDIRPATH_CODES")
			RunWait sDirPath . "\vbs\tools\win\file_ope\CopyAsWorkFile.vbs " . sTrgtPaths
			;FocusFileDirListAtExplorer()
		}
		^+g::	; Grep検索＠TresGrep
		{
			sTrgtPaths := GetCurDirPathAtExplorer()
			sExePath := EnvGet("MYEXEPATH_TRESGREP")
			Run sExePath . " " . sTrgtPaths
		}
		^+z::	; 圧縮/パスワード圧縮/解凍
		{
			global myGui
			global ogcListBoxAnswer
			myGui := Gui()
			myGui.OnEvent("Close", GuiEscape)
			myGui.OnEvent("Escape", GuiEscape)
			myGui.Add("Text", , "圧縮/パスワード圧縮/解凍を実行します。`n処理を選択してください。")
			ogcListBoxAnswer := myGui.Add("ListBox", "vAnswer Choose1 R3", ["圧縮", "パスワード圧縮", "解凍"])
			ogcButtonZipEnter := myGui.Add("Button", "Hidden w0 h0 Default", "ZipEnter")
			ogcButtonZipEnter.OnEvent("Click", ButtonZipEnter.Bind("Normal"))
			myGui.Show("Center")
		}
			ButtonZipEnter(A_GuiEvent, GuiCtrlObj, Info, *)
			{
				global myGui
				global ogcListBoxAnswer
				vAnswer := ogcListBoxAnswer.Text
				;MsgBox %vAnswer%
				myGui.Destroy()
				sTrgtPaths := GetSelFilePathAtExplorer(1)
				sDirPath := EnvGet("MYDIRPATH_CODES")
				If ( vAnswer == "圧縮" ) {
					RunWait(sDirPath . "\vbs\tools\7zip\ZipFile.vbs " . sTrgtPaths)
				} Else If ( vAnswer == "パスワード圧縮" ) {
					RunWait(sDirPath . "\vbs\tools\7zip\ZipPasswordFile.vbs " . sTrgtPaths)
				} Else If ( vAnswer == "解凍" ) {
					RunWait(sDirPath . "\vbs\tools\7zip\UnzipFile.vbs " . sTrgtPaths)
				} Else {
					MsgBox("`"[ERROR] 圧縮/パスワード圧縮/解凍 選択`"")
				}
				;FocusFileDirListAtExplorer()
				return
			}
		^s::	; ファイル作成
		{
			IB := InputBox("テキストファイルを作成します。`n処理を選択してください。", "", , ".txt"), sFileName := IB.Value, ErrorLevel := IB.Result="OK" ? 0 : IB.Result="CANCEL" ? 1 : IB.Result="Timeout" ? 2 : "ERROR"
			sDirPath := GetCurDirPathAtExplorer()
			Sleep(500)	; explorerのファイル選択ペインへの遷移待ち処理
			RunWait(A_ComSpec " /c copy nul " sFileName, sDirPath)
			;FocusFileDirListAtExplorer()
		}
		^+l::	; ショートカット/シンボリックリンク作成
		{
			global myGui
			global ogcListBoxAnswer
			myGui := Gui()
			myGui.Add("Text", , "ショートカット/シンボリックリンクを作成します。`n処理を選択してください。")
			ogcListBoxAnswer := myGui.Add("ListBox", "vAnswer Choose1 R2", ["ショートカット作成", "シンボリックリンク作成"])
			ogcButtonSelLinkEnter := myGui.Add("Button", "Hidden w0 h0 Default", "SelLinkEnter")
			ogcButtonSelLinkEnter.OnEvent("Click", ButtonSelLinkEnter.Bind("Normal"))
			myGui.Show("Center")
		}
			ButtonSelLinkEnter(A_GuiEvent, GuiCtrlObj, Info, *)
			{
				global myGui
				global ogcListBoxAnswer
				vAnswer := ogcListBoxAnswer.Text
				myGui.Destroy()
				sTrgtPaths := GetSelFilePathAtExplorer(1)
				sDirPath := EnvGet("MYDIRPATH_CODES")
				If ( vAnswer == "ショートカット作成" ) {
					RunWait(sDirPath . "\vbs\command\CreateShortcutFile.vbs " . sTrgtPaths . ".lnk " . sTrgtPaths)
				} Else If ( vAnswer == "シンボリックリンク作成" ) {
					RunWait(sDirPath . "\vbs\tools\win\file_ope\CreateSymbolicLink.vbs " . sTrgtPaths)
				} Else {
					MsgBox("`"[ERROR] ショートカット/シンボリックリンク作成`"")
				}
				;FocusFileDirListAtExplorer()
				return
			}
		^+r::	; リネーム用バッチファイル作成
		{
			sTrgtPaths := GetSelFilePathAtExplorer(1)
			sDirPath := EnvGet("MYDIRPATH_CODES")
			RunWait(sDirPath . "\vbs\tools\win\file_ope\CreateRenameBat.vbs " . sTrgtPaths)
			;FocusFileDirListAtExplorer()
			return
		}
		^+F3::	; 隠しファイル 表示非表示切替え
		{
			Send("!vhh")
			;FocusFileDirListAtExplorer()
		}
		^+F4::	; フォルダサイズ解析＠DiskInfo
		{
			sTrgtPaths := GetCurDirPathAtExplorer()
			sExePath := EnvGet("MYEXEPATH_DISKINFO3")
			StartProgramAndActivate( sExePath, sTrgtPaths )
		}
		^+F8::	; タグファイルを作成する
		{
			sTrgtPaths := GetCurDirPathAtExplorer()
			sDirPath := EnvGet("MYDIRPATH_CODES")
			RunWait(sDirPath . "\vbs\tools\ctags,gtags\CreateTagFiles.vbs " . sTrgtPaths)
			;FocusFileDirListAtExplorer()
		}
		^+F9::	; 配下全てをVimで開く
		{
			sTrgtPaths := GetCurDirPathAtExplorer()
			sDirPath := EnvGet("MYDIRPATH_CODES")
			Run(sDirPath . "\vbs\tools\vim\OpenAllFilesWithVim.vbs " . sTrgtPaths)
		}
		^+F10::	; コマンドプロンプトを開く
		{
			sDirPath := GetCurDirPathAtExplorer()
			Run(A_ComSpec " /k cd " sDirPath)
		}
		^+F11::	; パス一覧作成
		{
			global myGui
			global ogcListBoxAnswer
			myGui := Gui()
			myGui.Add("Text", , "パス一覧を作成します。`n処理を選択してください。")
			ogcListBoxAnswer := myGui.Add("ListBox", "vAnswer Choose1 R4", ["ファイル＆フォルダ一覧作成", "ファイル一覧作成", "フォルダ一覧作成", "フォルダツリー作成"])
			ogcButtonPathListEnter := myGui.Add("Button", "Hidden w0 h0 Default", "PathListEnter")
			ogcButtonPathListEnter.OnEvent("Click", ButtonPathListEnter.Bind("Normal"))
			myGui.Show("Center")
		}
			ButtonPathListEnter(A_GuiEvent, GuiCtrlObj, Info, *)
			{
				global myGui
				global ogcListBoxAnswer
				vAnswer := ogcListBoxAnswer.Text
				myGui.Destroy()
				sDirPath := GetCurDirPathAtExplorer()
				If ( vAnswer == "ファイル＆フォルダ一覧作成" ) {
					RunWait(A_ComSpec " /c dir /s /b /a > `"" sDirPath "\_PathList_FileDir.txt`"", sDirPath)
				} Else If ( vAnswer == "ファイル一覧作成" ) {
					RunWait(A_ComSpec " /c dir *.* /b /s /a:a-d > `"" sDirPath "\_PathList_File.txt`"", sDirPath)
				} Else If ( vAnswer == "フォルダ一覧作成" ) {
					RunWait(A_ComSpec " /c dir /b /s /a:d > `"" sDirPath "\_PathList_Dir.txt`"", sDirPath)
				} Else If ( vAnswer == "フォルダツリー作成" ) {
					RunWait(A_ComSpec " /c tree /f > `"" sDirPath "\_DirTree.txt`"", sDirPath)
				} Else {
					MsgBox("`"[ERROR] パス一覧作成`"")
				}
				;FocusFileDirListAtExplorer()
			}
		
		GuiEscape(*) {
		}
	;	GuiClose:
	;	ButtonCancel:
	;		myGui.Cancel()
	;		Return
	#HotIf
	
	#HotIf WinActive("ahk_exe EXCEL.EXE")
		F1::return	;F1ヘルプ無効化
		+Space::	;IME ON状態でShift+Space(行選択)が効かない対策
		{
			if (IME_GET() == 1) {
				IME_SET(0)
				Sleep 50
				SendInput "+{Space}"
				Sleep 50
				IME_SET(1)
			} else {
				SendInput "+{Space}"
			}
		}
	#HotIf
	
	#HotIf WinActive("ahk_exe iThoughts.exe")
		F1::return	;F1ヘルプ無効化
	#HotIf
	
	#HotIf WinActive("ahk_exe Rapture.exe")
		Esc::!F4	;Escで終了
	#HotIf
	
	#HotIf WinActive("ahk_exe vimrun.exe")
		Esc::!F4	;Escで終了
	#HotIf
	
	#HotIf WinActive("ahk_exe XF.exe")
		^WheelUp::SendInput "^+{Tab}"  ;Next tab.
		^WheelDown::SendInput "^{Tab}" ;Previous tab.
	#HotIf
	
	#HotIf WinActive("ahk_exe chrome.exe")
	;	^WheelUp::SendInput ^+{Tab}  ;Next tab.
	;	^WheelDown::SendInput ^{Tab} ;Previous tab.
	#HotIf
	
	#HotIf WinActive("ahk_class MPC-BE")
		]::Send "{Space}"
	#HotIf
	
	#HotIf WinActive("ahk_exe PDFXEdit.exe")
		MButton::	SendInput "^z" ;元に戻す
		XButton1::	SendInput "!5" ;下線
		XButton2::	SendInput "!4" ;テキストハイライト
	#HotIf

;* ***************************************************************
;* Functions
;* ***************************************************************
	StartProgramAndActivate( sExePath, sFilePath, bLaunchSingleProcess:=0 )
	{
		;*** preprocess ***
		If ( sExePath == "" or sFilePath == "" )
		{
			MsgBox "[ERROR] please specify arguments to StartProgramAndActivate()."
			return
		}
		sExeName := ExtractFileName(sExePath)
		sExeDirPath := ExtractDirPath(sExePath)
		sFileName := ExtractFileName(sFilePath)
		;MsgBox sExePath=%sExePath% `n sExeName=%sExeName% `n sExeDirPath=%sExeDirPath% `n sFilePath=%sFilePath%
		
		;*** check if the program is running ***
		If ( bLaunchSingleProcess == 1 ) {
			PID := ProcessExist(sExeName)
			If (PID == 0)
			{
				WinActivate "ahk_pid " . PID
				return
			}
		}
		
		;*** start program ***
		Try {
			Run sExePath . " " . sFilePath, sExeDirPath, , &sOutputVarPID
		} Catch Error {
			MsgBox "[error] run error : " . Error
		}
	;	WinActivate "ahk_pid " . sOutputVarPID
		return
	}
	
	; 起動＆アクティベート処理 (ファイルパス指定のみ)
	;
	; 備考：
	;   ・単一プロセス起動は指定不可。
	;       理由）単一プロセス起動は、プログラム名を基にしたプロセスの起動有無を
	;             確認することで実現できる。本関数はプログラム名を指定しないため、
	;             単一プロセス起動を実現できない。
	StartProgramAndActivateFile( sFilePath )
	{
		;*** preprocess ***
		If ( sFilePath == "" )
		{
			MsgBox "[ERROR] please specify arguments to StartProgramAndActivateFile()."
			return
		}
		sFileName := ExtractFileName(sFilePath)
		;MsgBox sFilePath=%sFilePath% `n sFileName=%sFileName%
		
		;*** start program ***
		Try {
			Run sFilePath, , , &sOutputVarPID
		} Catch Error {
			MsgBox "[error] run error : " . Error
		}
	;	WinActivate "ahk_pid " . sOutputVarPID
		return
	}
	
	; 起動＆アクティベート処理 (実行プログラム指定のみ)
	StartProgramAndActivateExe( sExePath, bLaunchSingleProcess:=0 )
	{
		;*** preprocess ***
		If ( sExePath == "" )
		{
			MsgBox "[ERROR] please specify arguments to StartProgramAndActivateExe()."
			return
		}
		
		sExeName := ExtractFileName(sExePath)
		sExeDirPath := ExtractDirPath(sExePath)
		;MsgBox sExePath=%sExePath% `n sExeDirPath=%sExeDirPath% `n sExeName=%sExeName%
		
		;*** check if the program is running ***
		If ( bLaunchSingleProcess == 1 ) {
			PID := ProcessExist(sExeName)
			If (PID != 0)
			{
				WinActivate "ahk_pid " . PID
				return
			}
		}
		
		;*** start program ***
		Try {
			Run sExePath, sExeDirPath, , &sOutputVarPID
		} Catch Error {
			MsgBox "[error] run error : " . Error
		}
	;	WinActivate "ahk_pid " . sOutputVarPID
		return
	}

	; 無変換キー同時押し実装
	MuhenkanSimultPush( sSendKey )
	{
		if(GetKeyState("Shift","P") and GetKeyState("Ctrl","P") and GetKeyState("Alt","P")){
			Send "!^+{" . sSendKey . "}"
		} else if(GetKeyState("Shift","P") and GetKeyState("Ctrl","P")){
			Send "^+{" . sSendKey . "}"
		} else if(GetKeyState("Shift","P") and GetKeyState("Alt","P")){
			Send "!+{" . sSendKey . "}"
		} else if(GetKeyState("Alt","P") and GetKeyState("Ctrl","P")){
			Send "!^{" . sSendKey . "}"
		} else if(GetKeyState("Alt","P")){
			Send "!{" . sSendKey . "}"
		} else if(GetKeyState("Ctrl","P")){
			Send "^{" . sSendKey . "}"
		} else if(GetKeyState("Shift","P")){
			Send "+{" . sSendKey . "}"
		} else {
			Send "{" . sSendKey . "}"
		}
		return
	}

	;Windowタイル切り替え
	GetWinTileModeMin()
	{
		iMonitorNum := SysGet(80) ; SM_CMONITORS: Number of display monitors on the desktop (not including "non-display pseudo-monitors").
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
			global giWinTileMode := iWinTileModeMin
		} else {
			global giWinTileMode := giWinTileMode + 1
		}
	;	MsgBox, giWinTileMode: %giWinTileMode%`n iWIN_TILE_MODE_MAX: %iWIN_TILE_MODE_MAX%`n iWinTileModeMin: %iWinTileModeMin%
	}
	DecrementWinTileMode()
	{
		iWinTileModeMin := GetWinTileModeMin()
		if ( giWinTileMode <= iWinTileModeMin ) {
			global giWinTileMode := iWIN_TILE_MODE_MAX
		} else {
			global giWinTileMode := giWinTileMode - 1
		}
	;	MsgBox, giWinTileMode: %giWinTileMode%`n iWIN_TILE_MODE_MAX: %iWIN_TILE_MODE_MAX%`n iWinTileModeMin: %iWinTileModeMin%
	}
	GetMonitorPosInfo( MonitorNum, &X, &Y, &Width, &Height )
	{
		try
		{
			ActualN := MonitorGetWorkArea(MonitorNum, &Left, &Top, &Right, &Bottom)
		;	MsgBox "Left: " Left " -- Top: " Top " -- Right: " Right " -- Bottom: " Bottom
		}
		catch
		{
			MsgBox "Monitor " . MonitorNum . " doesn't exist or an error occurred."
			return
		}
		Y := Top
		if ( Left < Right ) {
			X := Left
			Width := Right - Left + 1
		} else {
			X := Right
			Width := Left - Right + 1
		}
		Height := Bottom - Top + 1
	;	MsgBox, %MonitorNum%`n%X%`n%Y%`n%Width%`n%Height%
	}
	; ウィンドウサイズ切り替え
	ApplyWinTileMode()
	{
		GetMonitorPosInfo(1, &mainx, &mainy, &mainwidth, &mainheight )
		GetMonitorPosInfo(2, &subx, &suby, &subwidth, &subheight )
	;	MsgBox, mainx: %mainx%`nmainy: %mainy%`nmainwidth: %mainwidth%`nmainheight: %mainheight%`nsubx: %subx%`nsuby: %suby%`nsubwidth: %subwidth%`nsubheight: %subheight%
		
		winywhole := suby + ( subheight * iWIN_Y_OFFSET )
		winheightwhole := subheight * ( 1 - iWIN_Y_OFFSET )
	;	MsgBox, giWinTileMode: %giWinTileMode%`nwinywhole: %winywhole%`nwinheightwhole: %winheightwhole%
		if ( giWinTileMode = 0 ) {			;サブ全体
			winx := subx
			winwidth := subwidth
			winy := winywhole
			winheight := winheightwhole
		} else if ( giWinTileMode = 1 ) {	;サブ上
			winx := subx
			winwidth := subwidth
			winy := winywhole
			winheight := Integer(winheightwhole / 2)
		} else if ( giWinTileMode = 2 ) {	;サブ下
			winx := subx
			winwidth := subwidth
			winy := winywhole + Integer(winheightwhole / 2)
			winheight := Integer(winheightwhole / 2)
		} else if ( giWinTileMode = 3 ) {	;メイン全体
			winx := mainx
			winy := mainy
			winwidth := mainwidth
			winheight := mainheight
		} else if ( giWinTileMode = 4 ) {	;メイン左
			winx := mainx - iWIN_TILE_MODE_OFFSET
			winy := mainy
			winwidth := Integer(mainwidth / 2) + iWIN_TILE_MODE_OFFSET
			winheight := mainheight + iWIN_TILE_MODE_OFFSET
		} else if ( giWinTileMode = 5 ) {	;メイン右
			winx := mainx + Integer(mainwidth / 2) - iWIN_TILE_MODE_OFFSET
			winy := mainy
			winwidth := Integer(mainwidth / 2) + iWIN_TILE_MODE_OFFSET
			winheight := mainheight + iWIN_TILE_MODE_OFFSET
		} else {
			MsgBox "[error] invalid giWinTileMode.`n" . giWinTileMode
			return
		}
	;	MsgBox, giWinTileMode: %giWinTileMode%`nwinx: %winx%`nwiny: %winy%`nwinwidth: %winwidth%`nwinheight: %winheight%
		WinMove winx, winy, winwidth, winheight, "A"
		return
	}
	SetTimerWinTileMode()
	{
		SetTimer ClearWinTileMode, iWIN_TILE_MODE_CLEAR_INTERVAL
	}
	ClearWinTileMode()
	{
		global giWinTileMode
		giWinTileMode := iWIN_TILE_MODE_MAX
	;	TrayTip "タイルモードをクリアしました", "タイルモードクリアタイマー", 1
	;	Sleep 5000
	;	TrayTip
		Return
	}

	; ファイル名取得
	ExtractFileName( sFilePath )
	{
		SplitPath sFilePath, &sFileName, &sDirPath, &sExtName, &sFileBaseName, &sDrive
		sFileName := StrReplace(sFileName, "`"", )
	;	MsgBox sFilePath . "`n" . sFileName . "`n" . sDirPath . "`n" . sExtName . "`n" . sFileBaseName . "`n" . sDrive
		return sFileName
	}
	; ディレクトリパス取得
	ExtractDirPath( sFilePath )
	{
		SplitPath sFilePath, &sFileName, &sDirPath, &sExtName, &sFileBaseName, &sDrive
		sDirPath := StrReplace(sDirPath, "`"", )
	;	MsgBox sFilePath . "`n" . sFileName . "`n" . sDirPath . "`n" . sExtName . "`n" . sFileBaseName . "`n" . sDrive
		return sDirPath
	}

	; 選択ファイルパス取得＠explorer
	GetSelFilePathAtExplorer( bIsDelimiterSpace )
	{
		clipboard_old := A_Clipboard
		A_Clipboard := ""
		Send "!hcp"
		ClipWait
		sTrgtPaths := A_Clipboard
		A_Clipboard := clipboard_old
		if ( bIsDelimiterSpace = 1 ) {
			sTrgtPaths := StrReplace(sTrgtPaths, "`r`n", A_Space)
		}
	;	MsgBox sTrgtPaths=%sTrgtPaths%
		return sTrgtPaths
	}
	; 現在フォルダパス取得＠explorer
	GetCurDirPathAtExplorer()
	{
		clipboard_old := A_Clipboard
		A_Clipboard := ""
		Send "!d"
		Send "^c"
		ClipWait
	;	Send "{ESC}"
		Send "+{Tab 4}"
		sTrgtPaths := A_Clipboard
		A_Clipboard := clipboard_old
	;	MsgBox sTrgtPaths=%sTrgtPaths%
		return sTrgtPaths
	}
	; 選択ファイル名取得＠explorer
	GetSelFileNameAtExplorer()
	{
		sFilePaths := GetSelFilePathAtExplorer(0)
		sDirPaths := GetCurDirPathAtExplorer()
		sTrgtPaths := StrReplace(sFilePaths, sDirPaths . "\", )
		sTrgtPaths := StrReplace(sTrgtPaths, "`"", )
	;	MsgBox sTrgtPaths=%sTrgtPaths% `n sFilePaths=%sFilePaths% `n sDirPaths=%sDirPaths%
		return sTrgtPaths
	}
	; ファイルリストへフォーカスを移す＠explorer
	;FocusFileDirListAtExplorer()
	;{
	;	Sleep 100
	;	ControlFocus, SysTreeView321
	;	If ErrorLevel=0
	;	{
	;	;	MsgBox "ControlFocus success"
	;		Sleep 100
	;		Send "{Tab}"
	;	}
	;}
	FocusFileDirListAtExplorer()
	{
		WinActivate "ahk_class CabinetWClass ahk_exe Explorer.EXE"
		Sleep 200
		Send "^f"
		Sleep 200
		Send "{Tab}"
		Sleep 200
		Send "{Tab}"
		return
	}

	; IME.ahk
	; [URL] https://github.com/s-show/AutoHotKey/blob/AutoHotKey/IME.ahk
	;-----------------------------------------------------------
	; IMEの状態の取得
	;   WinTitle="A"    対象Window
	;   戻り値          1:ON / 0:OFF
	;-----------------------------------------------------------
	IME_GET(WinTitle:="A")  {
		Controls := WinGetControlsHwnd(WinTitle)
		hwnd := ControlGetHWND(Controls[1], WinTitle)
		if (WinActive(WinTitle))    {
			PtrSize := !A_PtrSize ? 4 : A_PtrSize
			stGTI := Buffer(cbSize:=4+4+(PtrSize*6)+16, 0) ; V1toV2: if 'stGTI' is a UTF-16 string, use 'VarSetStrCapacity(&stGTI, cbSize:=4+4+(PtrSize*6)+16)'
			NumPut "UInt", cbSize, stGTI   ;    DWORD   cbSize;
			hwnd := DllCall("GetGUIThreadInfo", "Uint", 0, "Ptr", stGTI)
					 ? NumGet(stGTI, 8+PtrSize, "UInt") : hwnd
		}
		return DllCall("SendMessage", "UInt", DllCall("imm32\ImmGetDefaultIMEWnd", "Uint", hwnd), "UInt", 0x0283, "Int", 0x0005, "Int", 0)
	}
	
	;-----------------------------------------------------------
	; IMEの状態をセット
	;	SetSts			1:ON / 0:OFF
	;	WinTitle="A"	対象Window
	;	戻り値			0:成功 / 0以外:失敗
	;-----------------------------------------------------------
	IME_SET(SetSts, WinTitle:="A")    {
		Controls := WinGetControlsHwnd(WinTitle)
		hwnd := ControlGetHWND(Controls[1], WinTitle)
		if (WinActive(WinTitle))    {
			PtrSize := !A_PtrSize ? 4 : A_PtrSize
			stGTI := Buffer(cbSize:=4+4+(PtrSize*6)+16, 0) ; V1toV2: if 'stGTI' is a UTF-16 string, use 'VarSetStrCapacity(&stGTI, cbSize:=4+4+(PtrSize*6)+16)'
			NumPut("UInt", cbSize, stGTI, 0)   ;    DWORD   cbSize;
			hwnd := DllCall("GetGUIThreadInfo", "Uint", 0, "Ptr", stGTI)
					 ? NumGet(stGTI, 8+PtrSize, "UInt") : hwnd
		}
		return DllCall("SendMessage", "UInt", DllCall("imm32\ImmGetDefaultIMEWnd", "Uint", hwnd), "UInt", 0x0283, "Int", 0x006, "Int", SetSts)
	}

	; DimMonitor.ahk
	; [URL] https://sites.google.com/site/bucuerider/autohotkey/dimmonitor
	DimMon_GenFilter()
	{
		global MonitorCount := MonitorGetCount()
		aDimGui := Array()
		global aDimId := Array()
	;	MsgBox "MonitorCount = " . MonitorCount
		global DimId
		Loop MonitorCount
		{
			MonitorGet(A_Index, &MonitorLeft, &MonitorTop, &MonitorRight, &MonitorBottom)
			Width := MonitorRight - MonitorLeft
			Height := MonitorBottom - MonitorTop
			aDimGui.push Gui()
			aDimGui[A_Index].Opt("+LastFound +ToolWindow -Disabled -SysMenu -Caption +E0x20 +AlwaysOnTop")
			aDimGui[A_Index].BackColor := "000000"	;フィルタの色（HTMLカラーコード参照）
			aDimGui[A_Index].Title := "DimMonitor" . A_Index
			aDimGui[A_Index].Show("X" . MonitorLeft . " Y" . MonitorTop . " W" . Width . " H" . Height)
			aDimId.push WinGetId("DimMonitor" . A_Index . " ahk_class AutoHotkeyGUI")
			DimId := aDimId[A_Index]
			WinSetTransparent(Integer(Dim * 255 / 100), "ahk_id " . DimId)
		;	MsgBox "MonitorCount = " . MonitorCount . ", A_Index = " . A_Index . ", DimId = " . DimId
		}
		Return
	}
	
	DimMon_HotKey()
	{
		DimMon_LoopMonitor()
		DimMon_AutoHideTip("明るさ：" . 100 - Dim . "%", 500)
		Return
	}
	
	DimMon_LoopMonitor()
	{
		global DimId
		global Dim
		global MonitorCount
		Loop MonitorCount
		{
			DimId := aDimId[A_Index]
			WinSetTransparent(Integer(Dim * 255 / 100), "ahk_id " . DimId)
		}
		Return
	}
	
	DimMon_AutoHideTip(Txt, Time)
	{
		ToolTip(Txt)
		SetTimer(DimMon_HideTip, -1 * Time)
		Return
	}
	
	DimMon_HideTip()
	{
		ToolTip()
		Return
	}
