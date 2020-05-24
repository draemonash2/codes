﻿	#NoEnv							; 通常、値が割り当てられていない変数名を参照しようとしたとき、システムの環境変数に同名の変数がないかを調べ、
	
									; もし存在すればその環境変数の値が参照される。スクリプト中に #NoEnv を記述することにより、この動作を無効化できる。
;	#Warn							; Enable warnings to assist with detecting common errors.
	SendMode Input					; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。
;	SetWorkingDir %A_ScriptDir%		; スクリプトの作業ディレクトリを本スクリプトの格納ディレクトリに変更。
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え

	#Include %A_ScriptDir%\lib\IME.ahk

;* ***************************************************************
;* Settings
;* ***************************************************************
DOC_DIR_PATH = C:\Users\%A_Username%\Dropbox\100_Documents

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

;***** Global *****
	;ホットキーヘルプ表示
		!^+F1::
			sFilePath = "C:\codes\ahk\UserDefHotKeyHotkeyHelp.vsdx"
			StartProgramAndActivate( "", sFilePath )
			return
	
	;#todo.itmz
		^+!Up::
			sExePath = "C:\Program Files (x86)\toketaWare\iThoughts\iThoughts.exe"
			sFilePath = "%DOC_DIR_PATH%\#todo.itmz"
			StartProgramAndActivate( sExePath, sFilePath )
			Send, {F2}{esc}
			return
	;#temp.txt
		^+!Down::
			sExePath = "C:\prg_exe\Vim\gvim.exe"
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
		^+!End::
			sFilePath = "%DOC_DIR_PATH%\210_【衣食住】家計\100_予算管理.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	
	;言語チートシート
		^+!_::
			sFilePath = "C:\codes\lang_cheet_sheet.xlsx"
			StartProgramAndActivate( "", sFilePath )
			return
	;$object.xlsm
		^+!?::
			sFilePath = "C:\other\template\$object.xlsm"
			StartProgramAndActivate( "", sFilePath )
			return
	
	;rapture.exe
		^+!x::
			Run "C:\prg_exe\Rapture\rapture.exe"
			return
	;KitchenTimer.vbs
		^+!k::
			Run "C:\codes\vbs\tools\win\other\KitchenTimer.vbs"
			return
	;xf.exe
		^+!z::
			Run "C:\prg_exe\X-Finder\XF.exe"
			return
	;cCalc.exe
		^+!;::
			RunSuppressMultiStart( "C:\prg_exe\cCalc\cCalc.exe", "" )
			return
	
	;Bluetoothテザリング起動
		^+!F10::
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
		^+!F12::
			sExePath = "C:\prg_exe\Vim\gvim.exe"
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
			
	;かなキーをコンテキストメニュー表示へ
		RAlt::AppsKey
			return
	;プリントスクリーン単押しを抑制
		PrintScreen::return
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

;***** Software local *****
	#IfWinActive ahk_exe EXCEL.EXE
		;F1ヘルプ無効化
			F1::return
		;Scroll left.
			+WheelUp::
			SetScrollLockState, On
			SendInput {Left}
			SetScrollLockState, Off
			Return
		;Scroll right.
			+WheelDown::
			SetScrollLockState, On
			SendInput {Right}
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
	
