﻿	#NoEnv							; 通常、値が割り当てられていない変数名を参照しようとしたとき、システムの環境変数に同名の変数がないかを調べ、
	
									; もし存在すればその環境変数の値が参照される。スクリプト中に #NoEnv を記述することにより、この動作を無効化できる。
;	#NoTrayIcon						; スクリプトのタスクトレイアイコンを非表示にする。
;	#Warn							; Enable warnings to assist with detecting common errors.
	SendMode Input					; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。
;	SetWorkingDir %A_ScriptDir%		; スクリプトの作業ディレクトリを本スクリプトの格納ディレクトリに変更。
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え
	#HotkeyModifierTimeout 100		; キーボードフックなしでホットキー中でSendコマンドを使用したときに修飾キーの状態を復元しなくなるタイムアウト時間を設定。

	#Include %A_ScriptDir%\lib\IME.ahk

;* ***************************************************************
;* Settings
;* ***************************************************************
DOC_DIR_PATH = C:\Users\%A_Username%\Dropbox\100_Documents

;* ***************************************************************
;* Timer
;* ***************************************************************
	SetTimer ClearWinTileMode, 5000
		Return
	ClearWinTileMode:
		;TrayTip, タイマーClearWinTileMode実行, iWinTileModeクリア, 1, 17
		iWinTileMode := 5
		Return

;* ***************************************************************
;* Keys
;* ***************************************************************
;[参考URL]
;	https://sites.google.com/site/autohotkeyjp/reference/KeyList
;		無変換）vk1Dsc07B
;		変換）	vk1Csc079
;		^）		Control
;		+）		Shift
;		!）		Alt
;		#）		Windowsロゴキー

;***** キー置き換え *****
	;無変換キー＋方向キーでPgUp,PgDn,Home,End
		vk1Dsc07B::vk1Dsc07B
		vk1Dsc07B & Right::MuhenkanSimultPush( "End" )
		vk1Dsc07B & Left::MuhenkanSimultPush( "Home" )
		vk1Dsc07B & Up::MuhenkanSimultPush( "PgUp" )
		vk1Dsc07B & Down::MuhenkanSimultPush( "PgDn" )

;***** ホットキー(Global) *****
	;ホットキー配置表示
		!^+F1::
			sFilePath = "C:\other\グローバルホットキー配置.vsdx"
			StartProgramAndActivate( "", sFilePath )
			return
	;ホットキーフォルダ表示
		!^+F12::
			sFilePath = "C:\Users\%A_Username%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\$Hotkey"
			StartProgramAndActivate( "", sFilePath )
			return
	
	;#todo.itmz
		^+!Up::
			EnvGet, sExePath, MYEXEPATH_ITHOUGHTS
			sFilePath = "%DOC_DIR_PATH%\#todo.itmz"
			StartProgramAndActivate( sExePath, sFilePath )
			Send, {F2}{esc}
			return
	;#temp.txt
		^+!Down::
			EnvGet, sExePath, MYEXEPATH_GVIM
			sFilePath = "%DOC_DIR_PATH%\#temp.txt"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	;#temp.xlsm
		^+!Right::
			sFilePath = "%DOC_DIR_PATH%\#temp.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	;#temp.vsdm
		^+!Left::
			sFilePath = "%DOC_DIR_PATH%\#temp.vsdm"
			StartProgramAndActivate( "", sFilePath )
			return
	
	;予算管理.xlsm
		^+!\::
			sFilePath = "%DOC_DIR_PATH%\210_【衣食住】家計\100_予算管理.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	
	;言語チートシート
		^+!c::
			sFilePath = "C:\other\言語チートシート.xlsx"
			StartProgramAndActivate( "", sFilePath )
			return
	;ショートカットキー
		^+!s::
			sFilePath = "C:\other\ショートカットキー一覧.xlsx"
			StartProgramAndActivate( "", sFilePath )
			return
	;$object.xlsm
		^+!o::
			sFilePath = "C:\other\template\#object.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	;用語集
		^+!/::
			sFilePath = "%DOC_DIR_PATH%\320_【自己啓発】勉強\words.itmz"
			StartProgramAndActivate( "", sFilePath )
			return
	;rapture.exe
		^+!x::
			EnvGet, sExePath, MYEXEPATH_RAPTURE
			Run %sExePath%
			return
	;KitchenTimer.vbs
		^+!k::
			EnvGet, sDirPath, MYDIRPATH_CODES
			Run % sDirPath . "\vbs\tools\win\other\KitchenTimer.vbs"
			return
	;xf.exe
		^+!z::
			EnvGet, sExePath, MYEXEPATH_XF
			Run %sExePath%
			return
	;cCalc.exe
		^+!;::
			EnvGet, sExePath, MYEXEPATH_CCALC
			RunSuppressMultiStart( sExePath, "" )
			return
	;Bluetoothテザリング起動
		^+!b::
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
			return
	;UserDefHotKey.ahk
		^+!a::
			EnvGet, sExePath, MYEXEPATH_GVIM
			sFilePath = "%A_ScriptFullPath%"
			StartProgramAndActivate( sExePath, sFilePath )
			return
	
	;Window最前面化
		Pause::
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
		global iWinTileMode := 0
		
		#LEFT::
			if iWinTileMode >= 5
			{
				iWinTileMode := 0
			}
			else
			{
				iWinTileMode++
			}
			ApplyWinTileMode( iWinTileMode )
			return
		#RIGHT::
			if iWinTileMode <= 0
			{
				iWinTileMode := 5
			}
			else
			{
				iWinTileMode--
			}
			ApplyWinTileMode( iWinTileMode )
			return
		ApplyWinTileMode( iWinTileMode )
		{
			;%s/[x]: /WinMove, A, , /g| %s/\v\t[ywh]: /, /g
			if iWinTileMode = 0
			{
				WinMove, A, , -2159, -2242, 2161, 3000	;サブ全体
			}
			else if iWinTileMode = 1
			{
				WinMove, A, , -2159, -2242, 2161, 1512	;サブ上
			}
			else if iWinTileMode = 2
			{
				WinMove, A, , -2159, -738, 2161, 1496	;サブ下
			}
			else if iWinTileMode = 3
			{
				WinMove, A, , 132, -8, 1796, 1096		;メイン全体
			}
			else if iWinTileMode = 4
			{
				WinMove, A, , 133, 0, 934, 1087			;メイン左
			}
			else
			{
				WinMove, A, , 1053, 0, 858, 1087		;メイン右
			}
			return
		}
	
	;プリントスクリーン単押しを抑制
		PrintScreen::return
	
;	;Teams一時退席抑止機能
;		+^!F11::
;			TrayTip, Teams一時退席抑止機能, Teamsの一時退席を抑止します。`nEscキー長押し(3秒以上)で停止できます。, 5, 17
;			Loop
;			{
;				Sleep, 3000
;				GetKeyState, sPressState, Esc, P
;				If sPressState = D
;				{
;					TrayTip, Teams一時退席抑止機能, Teamsの一時退席抑止を解除します。, 5, 17
;					Break
;				}
;				Else
;				{
;					Send, {vkF3sc029}
;				}
;			}
;			return
	
	;テスト用
		^Pause::
			MsgBox, ctrlpause
			Return
		+Pause::
			MsgBox, shiftpause
			Return
		+^!i::
			Send, ^c
			Sleep 200
			Send, ^e
			Sleep 200
			Send, ^v
			Sleep 200
			Send, {Left 3}
			Sleep 200
			Send, {Backspace 2}
			Sleep 200
			Send, {Space}
			Sleep 200
			Send, {End}
			Sleep 200
			Send, {Space}
			Sleep 200
			Send, openload
			return
	;	^1::
	;		MouseGetPos,x,y,hwnd,ctrl,3
	;		MouseClick, left, 1209, 932
	;		Sleep 100
	;		MouseClick, left, 1127, 1184
	;		Sleep 100
	;		MouseClick, left, 2089, 302
	;		Sleep 100
	;		MouseMove, x, y
	;		return

;***** ホットキー(Software local) *****
	;右Altキーをコンテキストメニュー表示に変更(WindowsTerminal以外)
	#IfWinNotActive ahk_exe WindowsTerminal.exe
		RAlt::
			Send, {AppsKey}
			return
	#IfWinNotActive
	
	#IfWinActive ahk_exe gimp-2.8.exe
		^Left::
			Send, {Left}{Backspace}{Esc}
			return
	#IfWinActive
	
	#IfWinActive ahk_exe EXCEL.EXE
		;F1ヘルプ無効化
			F1::return
		;Scroll left.
			+WheelUp::
			SetScrollLockState, On
			SendInput {Left 3}
			SetScrollLockState, Off
			Return
		;Scroll right.
			+WheelDown::
			SetScrollLockState, On
			SendInput {Right 3}
			SetScrollLockState, Off
			Return
		;Move prev sheet.
			^+WheelUp::
			SendInput ^{PgUp}
			Return
		;Move next sheet.
			^+WheelDown::
			SendInput ^{PgDn}
			Return
	#IfWinActive
	
	#IfWinActive ahk_exe iThoughts.exe
		;F1ヘルプ無効化
			F1::return
	#IfWinActive
	
	#IfWinActive ahk_exe Rapture.exe
		;Escで終了
			Esc::!F4
			return
	#IfWinActive
	
	#IfWinActive ahk_exe vimrun.exe
		;Escで終了
			Esc::!F4
			return
	#IfWinActive
	
	#IfWinActive AHK_Exe kinza.exe
		;The Great Suspender 用
			F8::^+s
			F9::^+u
			return
	#IfWinActive
	
	#IfWinActive ahk_class MPC-BE
			]::Send, {Space}
			return
	#IfWinActive
	
	;Kindle 自動ページ送り
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
	
	#IfWinActive ahk_exe PDFXCview.exe
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
	; 単一起動
	RunSuppressMultiStart( sExePath, sArguments )
	{
		IfInString, sExePath, \
		{
			Loop, Parse, sExePath , \
			{
				sExeName = %A_LoopField%
			}
			;MsgBox % sExeName
			Process, Exist, % sExeName
			If ErrorLevel<>0
			{
				WinActivate,ahk_pid %ErrorLevel%
			}
			else
			{
				Run % sExePath . " " . sArguments
			}
		}
		else
		{
			MsgBox sExePath
			MsgBox sArguments error!
		}
		return
	}
	
	; ★
	WinSizeChange( size, maxwinx, maxwiny )
	{
		if size = up
		{
			WinGetActiveStats, A, WinWidth, WinHeight, WinX, WinY
			if ( WinX = maxwinx && WinY = maxwiny )
			{
				WinMaximize
			}
			else
			{
				WinMaximize
			}
		}
		else if size = down
		{
			WinGetActiveStats, A, WinWidth, WinHeight, WinX, WinY
			if ( WinX = maxwinx && WinY = maxwiny )
			{
				WinRestore
			}
			else
			{
				WinMinimize
			}
		}
		else if size = max
		{
			WinMaximize
		}
		else if size = restore
		{
			WinRestore
		}
		else if size = min
		{
			WinMinimize
		}
		else
		{
			MsgBox "[error] please select up / down / max / restore / min."
		}
		return
	}
	
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
		IfInString, sFilePath, \
		{
			;*** extract file name ***
			;Loop, Parse, sFilePath , \
			;{
			;	sFileName = %A_LoopField%
			;}
			;StringReplace, sFileName, sFileName, ", , All
			Loop, Parse, sExePath , \
			{
				sExeName = %A_LoopField%
			}
			StringReplace, sExeName, sExeName, ", , All
			
			;*** for debug ***
			;MsgBox %sExePath%
			;MsgBox %sExeName%
			;MsgBox %sFilePath%
			;MsgBox %sFileName%
			
			;*** start program ***
			SetTitleMatchMode, 2 ;中間一致
			If ( sExePath == "" )
			{
				;MsgBox A ;for debug
				Run, %sFilePath%
			}
			else
			{
				;MsgBox B ;for debug
				Run, %sExePath% %sFilePath%
			}
			WinWait, ahk_exe %sExeName%, , 5
			If ErrorLevel <> 0
			{
				;MsgBox, could not be found %sExeName%.
				Return
			}
			
			;*** activate started program ***
			WinActivate, ahk_exe %sExeName%
			WinWaitActive, ahk_exe %sExeName%, , 5
			If ErrorLevel <> 0
			{
				;MsgBox, could not be activated %sExeName%.
				Return
			}
		}
		else
		{
			MsgBox sFilePath
			MsgBox argument error!
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
	
