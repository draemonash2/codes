;	#NoTrayIcon						; スクリプトのタスクトレイアイコンを非表示にする。
	#Warn							; Enable warnings to assist with detecting common errors.
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え
	#WinActivateForce				; ウィンドウのアクティブ化時に、穏やかな方法を試みるのを省略して常に強制的な方法でアクティブ化を行う。（タスクバーアイコンが点滅する現象が起こらなくなる）
	SendMode "Input"				; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。

;* ***************************************************************
;* Settings
;* ***************************************************************
DOC_DIR_PATH := "C:\Users\" . A_Username . "\Dropbox\100_Documents"
global iWIN_TILE_MODE_CLEAR_INTERVAL := 10000 ; [ms]
global iWIN_TILE_MODE_MAX := 3
global iWIN_Y_OFFSET := 2/7
global iWIN_TILE_MODE_OFFSET := 0
global bEnableAlwaysOnTop := 0

;* ***************************************************************
;* Define variables
;* ***************************************************************
global giWinTileMode := 0
global DimOld := 0
global Dim := 0
global DimId := 0

;* ***************************************************************
;* Preprocess
;* ***************************************************************
; TODO
;	SetTimer ClearWinTileMode, %iWIN_TILE_MODE_CLEAR_INTERVAL%
;		Return
;	ClearWinTileMode:
;		giWinTileMode := iWIN_TILE_MODE_MAX
;	;	TrayTip, タイマーClearWinTileMode実行, giWinTileMode = %giWinTileMode%, 1, 17
;		Return

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
; TODO
;	;無変換キー＋方向キーでPgUp,PgDn,Home,End
;		vk1Dsc07B::vk1Dsc07B
;		vk1Dsc07B & Right::	MuhenkanSimultPush( "End" )
;		vk1Dsc07B & Left::	MuhenkanSimultPush( "Home" )
;		vk1Dsc07B & Up::	MuhenkanSimultPush( "PgUp" )
;		vk1Dsc07B & Down::	MuhenkanSimultPush( "PgDn" )
;	;Insertキー
;		Insert::Return
;	;PrintScreenキー
;		PrintScreen::return

;***** ホットキー(Global) *****
	;UserDefHotKey.ahk
		^+!a::
		{
			sExePath := EnvGet("MYEXEPATH_GVIM")
			sFilePath := A_ScriptFullPath
			Run sExePath . " " . sFilePath
		}
	;ホットキー配置表示
		!^+F1::
		{
			sFilePath := "C:\other\グローバルホットキー配置.vsdx"
			Run sFilePath
		}
	;Programsフォルダ表示
		!^+F12::
		{
			sFilePath := "C:\Users\" . A_Username . "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
			Run sFilePath
			Sleep 100
			Send "+{tab}"
		}
	;#todo.itmz
		^+!Up::
		{
			sFilePath := DOC_DIR_PATH . "\#todo.itmz"
		;	lPID := ProcessWait("Dropbox.exe", 30) ; Dropboxが起動(≒同期が完了)するまで待つ(タイムアウト時間30s)
			Run sFilePath
		}
	;#temp.txt
		^+!Down::
		{
			sFilePath := "C:\Users\draem\Dropbox\100_Documents\#temp.txt"
			Run sFilePath
		}
		^+!Space::
		{
			sFilePath := "C:\Users\draem\Dropbox\100_Documents\#temp.txt"
			Run sFilePath
		}
	;#temp.xlsm
		^+!Right::
		{
			sFilePath := DOC_DIR_PATH . "\#temp.xlsm"
			Run sFilePath
		}
	;#temp.vsdm
		^+!Left::
		{
			sFilePath := DOC_DIR_PATH . "\#temp.vsdm"
			Run sFilePath
		}
	;予算管理.xlsm
		^+!\::
		{
			sFilePath := DOC_DIR_PATH . "\210_【衣食住】家計\100_予算管理.xlsm"
			Run sFilePath
		}
	;予算管理＠家族用.xlsx
		^+!^::
		{
			sFilePath := DOC_DIR_PATH . "\..\000_Public\家計\予算管理＠家族用.xlsx"
			Run sFilePath
		}
	;言語チートシート
		^+!c::
		{
			sFilePath := "C:\other\言語チートシート.xlsx"
			Run sFilePath
		}
	;ショートカットキー
		^+!s::
		{
			sFilePath := "C:\other\ショートカットキー一覧.xlsx"
			Run sFilePath
		}
	;#object.xlsm
		^+!o::
		{
			sFilePath := "C:\other\template\#object.xlsm"
			Run sFilePath
		}
	;用語集
		^+!/::
		{
			sFilePath := DOC_DIR_PATH . "\320_【自己啓発】勉強\words.itmz"
			Run sFilePath
		}
	;codes同期
		^+!y::
		{
			sDirPath := EnvGet("MYDIRPATH_CODES")
			sFilePath := sDirPath . "\_sync_github-codes-remote.bat"
			Run sFilePath
		}
	;KitchenTimer.vbs
		^+!k::
		{
			sDirPath := EnvGet("MYDIRPATH_CODES")
			sFilePath := sDirPath . "\vbs\tools\win\other\KitchenTimer.vbs"
			Run sFilePath
		}
	;定期キー送信
		^+!t::
		{
			sDirPath := EnvGet("MYDIRPATH_CODES")
			sFilePath := sDirPath . "\vbs\tools\win\other\PeriodicKeyTransmission.bat"
			Run sFilePath
		}
	;rapture.exe
		^+!x::
		{
			; TODO
;			; 明るさを最大にする
;			DimOld := Dim
;			Dim := 0
;			GoSub, LoopMonitor
			; Rapture 起動
			sExePath := EnvGet("MYEXEPATH_RAPTURE")
			Run sExePath
;			; 明るさを元に戻す
;			Sleep 5000
;			Dim := DimOld
;			GoSub, LoopMonitor
		}
	;xf.exe
	/*
		^+!z::
		{
			sExePath := EnvGet("MYEXEPATH_XF")
			Run sExePath ; TODO: single instance
			
		}
	*/
	;DOC_DIR_PATHフォルダ表示
		!^+z::
		{
			sFilePath := DOC_DIR_PATH
			Run sFilePath
			Sleep 100
			Send "+{tab}"
		}
	;cCalc.exe
		^+!;::
		{
			; TODO: Path to cCalc.dat
			sExePath := EnvGet("MYEXEPATH_CALC")
			Run sExePath ; TODO: single instance
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
		; TODO: test
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
		; TODO: test
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
			IncrementWinTileMode()
			ApplyWinTileMode()
		}
		!#RIGHT::
		{
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
		; ★custum mod <TOP>
			Dim += 20
		; ★custum mod <END>
			if (Dim > 80)
				Dim := 80
			DimMon_HotKey()
		}
		
		#PgUp::							; 明度を上げる（不透明度を下げる）
		{
			global Dim
		; ★custum mod <TOP>
			Dim -= 20
		; ★custum mod <END>
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
			sFilePath := "C:\Users\draem\Dropbox\100_Documents\#temp.txt"
			sRet := ExtractDirPath(sFilePath)
			MsgBox sRet
		}

;***** ホットキー(Software local) *****
	#HotIf !WinActive("ahk_exe WindowsTerminal.exe")
		RAlt::Send "{AppsKey}"	;右Altキーをコンテキストメニュー表示に変更
	#HotIf
	
	#HotIf WinActive("ahk_exe explorer.exe")
	; TODO:
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
			ActualN := MonitorGet(MonitorNum, &Left, &Top, &Right, &Bottom)
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

; TODO: Implement
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
		sTrgtPaths := StrReplace(sTrgtPaths, sDirPaths . "\", A_Space)
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
