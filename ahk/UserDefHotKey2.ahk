; [Help] https://www.autohotkey.com/docs/v2/index.htm

;	#NoTrayIcon						; スクリプトのタスクトレイアイコンを非表示にする。
	#Warn All						; Enable warnings to assist with detecting common errors.
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え
	#WinActivateForce				; ウィンドウのアクティブ化時に、穏やかな方法を試みるのを省略して常に強制的な方法でアクティブ化を行う。（タスクバーアイコンが点滅する現象が起こらなくなる）
	SendMode "Input"				; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。

;* ***************************************************************
;* Setting value
;* ***************************************************************
; {{{
global gsDOC_DIR_PATH := "C:\Users\" . A_Username . "\Dropbox\100_Documents"
global gsUSER_PROFILE_PATH := EnvGet("USERPROFILE")
global gsCONFIG_DIR_NAME := "UserDefHotKey"
global giWIN_TILE_MODE_CLEAR_INTERVAL_MS := 10000
global giWIN_TILE_MODE_RANGE_MIN := 1 ; 1～6 (Mon1:1-3, Mon2:4-6)
global giWIN_TILE_MODE_RANGE_MAX := 4 ; 1～6 (Mon1:1-3, Mon2:4-6)
global giWIN_TILE_MODE_WIN_RANGE_RATE := 5/7 ; 0～1
global giWIN_TILE_MODE_INIT := 0
global giSCREEN_BRIGHTNESS_STEP := 10 ; 0～100 [%]
global giSCREEN_BRIGHTNESS_MIN := giSCREEN_BRIGHTNESS_STEP ; 0～100 [%]
global giSCREEN_BRIGHTNESS_MAX := 100 ; 0～100 [%]
global giSCREEN_BRIGHTNESS_INIT_DAY := giSCREEN_BRIGHTNESS_MAX
global giSCREEN_BRIGHTNESS_INIT_NIGHT := 30
global giSCREEN_BRIGHTNESS_DAY_START_TIME := 7
global giSCREEN_BRIGHTNESS_DAY_END_TIME := 20
global giSTART_PRG_TOOLTIP_SHOW_TIME_MS := 2000
global giSLEEPPREVENT_INTERVAL_TIME_MS := 120000
global giSLEEPPREVENT_EXE_NAME := "javaw.exe" ; TurboVNC
global giSLEEPPREVENT_PROGRAM_NAME := "TurboVNC"
global giSLEEPPREVENT_KEY_NAME := " "
global gbSLEEPPREVENT_SHOW_TRAYTIP_WITH_ACT := False
global giJUMPCURSOL_KEYPRESS_NUM := 3
global giMOVECURSOL_MOVE_OFFSET_NEAR := 50
global giMOVECURSOL_MOVE_OFFSET_FAR := 150
global gbRALT2APPSKEY_RALT_TO_APPSKEY := false
global gfKITCHENTIMER_INIT_MIN := 3
global gbKITCHENTIMER_SAVE_INIT_MIN := true
global giKITCHENTIMER_TRAYTIP_DURATION_MS := 5000
global gsKITCHENTIMER_CONFIG_FILE_NAME := "KitchenTimer.cfg"
global giALARMTIMER_TRAYTIP_DURATION_MS := 5000
global gbALARMTIMER_INITTIME_CUR := false
global giALARMTIMER_INITTIME_MIN_STEP := 30 ;「0より大きい」「60以下」「60の約数である」をすべて満たす必要がある
global aiALARMTIMER_EVERYDAY_TRGT_WEEKDAY := [2, 3, 4, 5, 6] ; 1:Sun, 2:Mon, ... 7:Sat
global giALARMTIMER_SNOOZE_MSG_DURATION_SEC := 10
; }}}

;* ***************************************************************
;* Preprocess
;* ***************************************************************
; {{{
TraySetIcon "UserDefHotKey2.ico"
ShowAutoHideTrayTip("", A_ScriptName . " is loaded.", 2000)
StoreCurYearMonths()
InitScreenBrightness()
InitWinTileMode()
InitSleepPreventing()
;InitRAltAppsKeyMode()
RestartAlermTimer()
RestartKitchenTimer()
SetEveryDayAlermTimer()
; }}}

;* ***************************************************************
;* Keys
;*  [参考URL]
;*		https://www.autohotkey.com/docs/v2/KeyList.htm
;*			^）		Control
;*			+）		Shift
;*			!）		Alt
;*			#）		Windowsロゴキー
;*  [備考]
;*		Pause … HP製PC以外) Alt+Pause(Fn＋Shift)、HP製PC) Shift+Alt+Fn
;* ***************************************************************

;***** キー置き換え *****
	;無変換/変換キー単押し ; {{{
		VK1D::VK1D
		VK1C::VK1C
	; }}}
	;変換キー＋wasd → マウスカーソル移動 ; {{{
		VK1C & w::		MoveCursor("Up")
		VK1C & s::		MoveCursor("Down")
		VK1C & d::		MoveCursor("Right")
		VK1C & a::		MoveCursor("Left")
	; }}}
	;変換キー＋Space → マウスクリック ; {{{
		VK1C & Space::
		{
			if (GetKeyState("Shift","P")) {
				Click "Right"
			} else {
				Click
			}
		}
	; }}}
	;無変換キー＋方向キー → PgUp,PgDn,Home,End ; {{{
		; e.g. 無変換+上キー -> PgUp
		; e.g. 無変換+Shift+Alt+上キー -> Shift+Alt+PgUp
		VK1D & Right::	SendKeyWithModKeyCurPressing( "End" )
		VK1D & Left::	SendKeyWithModKeyCurPressing( "Home" )
		VK1D & Up::		SendKeyWithModKeyCurPressing( "PgUp" )
		VK1D & Down::	SendKeyWithModKeyCurPressing( "PgDn" )
	; }}}
	;無変換キー＋jkhl → 矢印キー ; {{{
		VK1D::VK1D
		VK1D & k::		Send "{Up}"
		VK1D & j::		Send "{Down}"
		VK1D & l::		Send "{Right}"
		VK1D & h::		Send "{Left}"
	; }}}
	;キー無効化 ; {{{
		Insert::Return																												;Insertキー
		PrintScreen::return																											;PrintScreenキー
	; }}}

;***** ホットキー（Global） *****
	;スクリプトリロード ; {{{
		^+!F5::		ReloadMe()
	; }}}
	;ファイルオープン ; {{{
		^+!a::		StartProgramAndActivate( EnvGet("MYEXEPATH_GVIM"), A_ScriptFullPath )											;UserDefHotKey.ahk
		^+!Down::	StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#temp.txt" )													;#temp.txt
		^+!Up::																														;#todo.itmz
		{
		;	lPID := ProcessWait("Dropbox.exe", 30) ; Dropboxが起動(≒同期が完了)するまで待つ(タイムアウト時間30s)
			StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#todo.itmz" )
		}
		^+!Right::	StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#temp.xlsm" )													;#temp.xlsm
		^+!Left::	StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#temp.vsdm" )													;#temp.vsdm
		^+!\::		StartProgramAndActivateFile( gsDOC_DIR_PATH . "\210_【衣食住】家計\100_予算管理.xlsm" )							;予算管理.xlsm
		^+!^::		StartProgramAndActivateFile( gsDOC_DIR_PATH . "\..\000_Public\家計\予算管理＠家族用.xlsx" )						;予算管理＠家族用.xlsx
		^+!/::		StartProgramAndActivateFile( gsDOC_DIR_PATH . "\320_【自己啓発】勉強\words.itmz" )								;用語集
		^+!o::		StartProgramAndActivateFile( "C:\other\template\#object.xlsm" )													;#object.xlsm
		^+!c::		StartProgramAndActivateFile( "C:\other\言語チートシート.xlsx" )													;言語チートシート
		^+!s::		StartProgramAndActivateFile( "C:\other\ショートカットキー一覧.xlsx" )											;ショートカットキー一覧
		^+!m::		StartProgramAndActivateFile( "C:\other\PC移行時チェックリスト.xlsx" )											;PC移行時チェックリスト.xlsx
	; }}}
	;仕事用 ; {{{
		^+!Space::	StartProgramAndActivateFile( gsUSER_PROFILE_PATH . "\_root\#temp.txt" )											;#temp.txt
		^+!Enter::	StartProgramAndActivateFile( gsUSER_PROFILE_PATH . "\_root\#memo.xlsm" )										;#memo.xlsm
		^+!@::		StartProgramAndActivateFile( gsUSER_PROFILE_PATH . "\_root\#memo_image.xlsx" )									;#memo_image.xlsx
		^+!-::		StartProgramAndActivateFile( gsUSER_PROFILE_PATH . "\_root\#timemng.xlsm" )										;#timemng.xlsm
		^+!0::		StartProgramAndActivateFile( gsUSER_PROFILE_PATH . "\_root\10_workitem\230901_教育_キャッチアップ\#memo_キャッチアップ.xlsm" )
		^+!9::		StartProgramAndActivateFile( gsUSER_PROFILE_PATH . "\_root\10_workitem\230922_開発_シミュレーション環境構築\#memo_シミュレーション環境構築.xlsm" )
		^+!8::		StartProgramAndActivateFile( gsUSER_PROFILE_PATH . "\_root\10_workitem\230922_開発_シミュレーション環境構築\20_output\231031_gen_world\v0.4以降\design_memo.xlsx" )
		^+!7::		StartProgramAndActivateFile( gsUSER_PROFILE_PATH . "\_root\10_workitem\230922_開発_シミュレーション環境構築\20_output\240628_auto_sim_test\logic_memo.txt" )
	; }}}
	;プログラム起動 ; {{{
		^+!y::		StartProgramAndActivateFile( EnvGet("MYDIRPATH_CODES") . "\_sync_github-codes-remote.bat" )						;codes同期
	;	^+!k::		StartProgramAndActivateFile( EnvGet("MYDIRPATH_CODES") . "\vbs\tools\win\other\KitchenTimer.vbs" )				;KitchenTimer.vbs
	;	^+!k::		Run A_ComSpec . " /c start ms-clock:"																			;クロックアプリ
		^+!k::		SetKitchenTimer()																								;キッチンタイマー
		^+!r::		SetAlermTimer()																									;アラームタイマー
		^+!t::		StartProgramAndActivateFile( EnvGet("MYDIRPATH_CODES") . "\vbs\tools\win\other\PeriodicKeyTransmission.bat" )	;定期キー送信
		^+!w::		StartProgramAndActivateFile( EnvGet("MYDIRPATH_CODES") . "\vbs\tools\win\file_ope\CopyRefFileFromWeb.vbs" )		;Webから参照ファイル取得
	;	^+!;::		StartProgramAndActivateExe( EnvGet("MYEXEPATH_CALC"), True )													;cCalc.exe
		^+!;::		Run A_ComSpec . " /c calc"																						;電卓アプリ
		^+!x::																														;rapture.exe
		{
			SetBrightnessTemporary(giSCREEN_BRIGHTNESS_MAX, 5000)
			StartProgramAndActivateExe( EnvGet("MYEXEPATH_RAPTURE"), False, False )
		}
	; }}}
	;フォルダ表示 ; {{{
		^+!z::																														;ファイラ―
		{
			;xf.exe
			StartProgramAndActivateExe( EnvGet("MYEXEPATH_XF"), 1 )
		;	;エクスプローラー
		;	StartProgramAndActivateFile( gsDOC_DIR_PATH )
		;	Sleep 100
		;	Send "+{tab}"
		}
		^+!F12::																													;Programsフォルダ表示
		{
			StartProgramAndActivateFile( "C:\Users\" . A_Username . "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs" )
			Sleep 100
			Send "+{tab}"
		}
	; }}}
	;サイトオープン ; {{{
		^+!1::	Run "https://draemonash2.github.io/"																				;Github.io
		^+!2::	Run "https://draemonash2.github.io/linux_os/linux.html"																;Github.io linux
		^+!3::	Run "https://draemonash2.github.io/gitcommand_lng/gitcommand.html"													;Github.io git command
		^+!h::	Run "https://www.deepl.com//translator"																				;翻訳サイト
	; }}}
	;Wifi接続 ; {{{
		/*
		^+!F9::																														;Bluetoothテザリング起動
		{
			Run "control printers"
			
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
		}
		^+!F9::	Run EnvGet("MYDIRPATH_CODES") . "\bat\tools\other\ConnectWifi.bat MyPerfectiPhone"									; Wifiテザリング
		*/
	; }}}
	;Windowタイル切り替え ; {{{
		!#LEFT::
		{
			;MsgBox "!#LEFT"
			SetTimerClearWinTileMode()
			IncrementWinTileMode()
			ApplyWinTileMode()
		}
		!#RIGHT::
		{
			;MsgBox "!#RIGHT"
			SetTimerClearWinTileMode()
			DecrementWinTileMode()
			ApplyWinTileMode()
		}
	; }}}
	;画面明るさ設定 ; {{{
		#Home::	SetBrightness(giSCREEN_BRIGHTNESS_MAX)
		#End::	SetBrightness(giSCREEN_BRIGHTNESS_MIN)
		#PgDn::	DarkenScreen()
		#PgUp::	BrightenScreen()
	; }}}
	;その他 ; {{{
	;	^+!r::		SetSleepPreventingMode("Toggle", True)																			;TurboVNCスリープ抑制
		!Pause::	ToggleAlwaysOnTopEnable()																						;Window最前面化
	;	^+!F11::	SwitchRAltAppsKeyMode()																							;右Alt->AppsKey置換え切替え
		Ctrl::																														;モニタ中心にカーソル移動
		{
			Loop giJUMPCURSOL_KEYPRESS_NUM - 1
			{
				Sleep 100
				iIsTimeout := KeyWait("Ctrl", "D T0.2")
				If (iIsTimeout == 0){
					Return
				}
			}
			;MsgBox "Ctrlキーが" . giJUMPCURSOL_KEYPRESS_NUM . "回押されました"
			MoveCursolToMonitorCenter()
		}
	; }}}
	;テスト用 ; {{{
		/*
		^Pause::	MsgBox "ctrlpause"
		+Pause::	MsgBox "shiftpause"
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
		/*
		^1::
		{
			;sImgFilePath := "C:\Users\draem\OneDrive\デスクトップ\schedule.jpg"
			;sOptions := "W300 H300"
			MyGui := Gui("+Resize")
			MyBtn := MyGui.Add("Button", "default", "&Load New Image")
			MyBtn.OnEvent("Click", LoadNewImage)
			MyRadio := MyGui.Add("Radio", "ym+5 x+10 checked", "Load &actual size")
			MyGui.Add("Radio", "ym+5 x+10", "Load to &fit screen")
			MyPic := MyGui.Add("Pic", "xm")
			MyGui.Show()
			LoadNewImage(*)
			{
				;Image := FileSelect(,, "Select an image:", "Images (*.gif; *.jpg; *.bmp; *.png; *.tif; *.ico; *.cur; *.ani; *.exe; *.dll)")
				;if Image = ""
				;	return
				Image := "C:\Users\draem\OneDrive\デスクトップ\schedule.jpg"
				if (MyRadio.Value)  ; Display image at its actual size.
				{
					Width := 0
					Height := 0
				}
				else ; Second radio is selected: Resize the image to fit the screen.
				{
					Width := A_ScreenWidth - 28  ; Minus 28 to allow room for borders and margins inside.
					Height := -1  ; "Keep aspect ratio" seems best.
				}
				MyPic.Value := Format("*w{1} *h{2} {3}", Width, Height, Image)  ; Load the image.
				MyGui.Title := Image
				MyGui.Show("xCenter y0 AutoSize")  ; Resize the window to match the picture size.
			}
		}
		*/
	; }}}

;***** ホットキー(Software local) *****
	#HotIf !WinActive("ahk_exe WindowsTerminal.exe") ; {{{
	;	RAlt::PressRAlt()		;右Altキーをコンテキストメニュー表示に変更
		#Hotif gbRALT2APPSKEY_RALT_TO_APPSKEY
			RAlt::Send "{AppsKey}"	;右Altキーをコンテキストメニュー表示に変更
		#HotIf
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe msedge.exe") ; {{{
		^!t::	;タブを複製して和訳
		{
			;タブを複製
			SendInput "^+k"
			sleep 1000
			;和訳
			SendInput "{AppsKey}"
			sleep 500
			SendInput "+t"
		}
		~RButton & WheelUp::SendInput "^+{Tab}"
		~RButton & WheelDown::SendInput "^{Tab}"
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe explorer.exe") ; {{{
		+F1::	Run EnvGet("MYDIRPATH_CODES") . "\vbs\tools\wimmerge\CompareWithWinmerge.vbs " . GetSelFilePathAtExplorer(1)		; winmergeで開く
		+F2::	StartProgramAndActivate( EnvGet("MYEXEPATH_GVIM"), GetSelFilePathAtExplorer(1) )									; vimで開く
		+F3::	StartProgramAndActivate( EnvGet("MYEXEPATH_VSCODE"), GetSelFilePathAtExplorer(1) )									; VSCodeで開く
		+F4::	StartProgramAndActivate( EnvGet("MYEXEPATH_HIDEMARU"), GetSelFilePathAtExplorer(1) )								; 秀丸で開く
		+F5::	StartProgramAndActivate( EnvGet("MYEXEPATH_EXCEL"), GetSelFilePathAtExplorer(1) )									; EXCELで開く
		+F9::	RunWait EnvGet("MYDIRPATH_CODES") . "\vbs\tools\win\file_ope\CopyAsWorkFile.vbs " . GetSelFilePathAtExplorer(1)		; 作業ファイルとしてコピー
		^+F3::	Send("!vhh")																										; 隠しファイル 表示非表示切替え
		^+F4::	StartProgramAndActivate( EnvGet("MYEXEPATH_DISKINFO3"), GetCurDirPathAtExplorer() )									; フォルダサイズ解析＠DiskInfo
		^+F8::	RunWait EnvGet("MYDIRPATH_CODES") . "\vbs\tools\ctags,gtags\CreateTagFiles.vbs " . GetCurDirPathAtExplorer()		; タグファイルを作成する
		^+F9::	Run EnvGet("MYDIRPATH_CODES") . "\vbs\tools\vim\OpenAllFilesWithVim.vbs " . GetCurDirPathAtExplorer()				; 配下全てをVimで開く
		^+F10::	Run A_ComSpec . " /k cd " . GetCurDirPathAtExplorer()																; コマンドプロンプトを開く
		^+F11::	CreateSlctCmndWindowPathList()																						; パス一覧作成
		^+c::	SetClipboard(GetSelFilePathAtExplorer(0))																			; ファイルパスコピー
		^+d::	SetClipboard(GetSelFileNameAtExplorer())																			; ファイル名コピー
		^+g::	StartProgramAndActivate( EnvGet("MYEXEPATH_TRESGREP"), GetCurDirPathAtExplorer() )									; Grep検索＠TresGrep
		^+z::	CreateSlctCmndWindowZip()																							; 圧縮/パスワード圧縮/解凍
		^+l::	CreateSlctCmndWindowLink()																							; ショートカット/シンボリックリンク作成
		^+r::	RunWait EnvGet("MYDIRPATH_CODES") . "\vbs\tools\win\file_ope\CreateRenameBat.vbs " . GetSelFilePathAtExplorer(1)	; リネーム用バッチファイル作成
		^s::																														; ファイル作成
		{
			IB := InputBox("テキストファイルを作成します。`n処理を選択してください。", "", , ".txt"), sFileName := IB.Value, ErrorLevel := IB.Result="OK" ? 0 : IB.Result="CANCEL" ? 1 : IB.Result="Timeout" ? 2 : "ERROR"
			Sleep(500)	; explorerのファイル選択ペインへの遷移待ち処理
			RunWait A_ComSpec " /c copy nul " sFileName, GetCurDirPathAtExplorer()
			;FocusFileDirListAtExplorer()
		}
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe EXCEL.EXE") ; {{{
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
	;	;Ctrl+Shift+ホイールUp/Downで右/左スクロール（旧Excel用）
	;	^+WheelUp::
	;	{
	;		SetScrollLockState True
	;		SendInput "{Left 5}"
	;		SetScrollLockState False
	;	}
	;	^+WheelDown::
	;	{
	;		SetScrollLockState True
	;		SendInput "{Right 5}"
	;		SetScrollLockState False
	;	}
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe iThoughts.exe") ; {{{
		F1::return	;F1ヘルプ無効化
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe Rapture.exe") ; {{{
		Esc::!F4	;Escで終了
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe vimrun.exe") ; {{{
		Esc::!F4	;Escで終了
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe XF.exe") ; {{{
		^WheelUp::SendInput "^+{Tab}"  ;Next tab.
		^WheelDown::SendInput "^{Tab}" ;Previous tab.
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe chrome.exe") ; {{{
	;	^WheelUp::SendInput ^+{Tab}  ;Next tab.
	;	^WheelDown::SendInput ^{Tab} ;Previous tab.
	#HotIf ; }}}
	#HotIf WinActive("ahk_class MPC-BE") ; {{{
		]::Send "{Space}"
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe PDFXEdit.exe") ; {{{
		MButton::	SendInput "^z" ;元に戻す
		XButton1::	SendInput "!5" ;下線
		XButton2::	SendInput "!4" ;テキストハイライト
	#HotIf ; }}}
	#HotIf WinActive("ahk_exe java.exe") and WinActive("TurboVNC: ") ; {{{
		; 特定位置へカーソル移動
		;   UbuntuのターミナルからGUIプログラムを起動後、
		;   自動的にターミナルにフォーカスを戻すために用意したマクロ
		VK1C & v::
		{
			MouseMove 2000, 550
		;	sleep 1000
			Click
		}
	#HotIf ; }}}

;* ***************************************************************
;* Functions (macro)
;* ***************************************************************
	; 起動＆アクティベート処理 (実行プログラム＆ファイルパス指定)
	StartProgramAndActivate( sExePath, sFilePath, bLaunchSingleProcess:=False, bShowToolTip:=True ) ; {{{
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
		;MsgBox "sExePath = " . sExePath . "`nsExeName = " . sExeName . "`nsExeDirPath = " . sExeDirPath . "`nsFilePath = sFilePath" . "`nbLaunchSingleProcess = " . bLaunchSingleProcess
		
		;*** show tooltip ***
		If ( bShowToolTip == True ) {
			ShowAutoHideToolTip(sFileName . " is starting...", giSTART_PRG_TOOLTIP_SHOW_TIME_MS)
		}
		
		;*** check if the program is running ***
		If ( bLaunchSingleProcess == True ) {
			iPID := ProcessExist(sExeName)
			If (iPID != 0)
			{
				WinActivate "ahk_pid " . iPID
				return
			}
		}
		
		;*** start program ***
		Try {
			Run sExePath . " " . sFilePath, sExeDirPath, , &sOutputVarPID
			WinWait("ahk_pid " . sOutputVarPID, , 1)
			WinActivate "ahk_pid " . sOutputVarPID
		;	BringActiveWindowToTop()
		} Catch Error as err {
		;	MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
		;		, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
		}
		ToolTip()
		return
	} ; }}}
	
	; 起動＆アクティベート処理 (ファイルパス指定のみ)
	;
	; 備考：
	;   ・単一プロセス起動は指定不可。
	;       理由）単一プロセス起動は、プログラム名を基にしたプロセスの起動有無を
	;             確認することで実現できる。本関数はプログラム名を指定しないため、
	;             単一プロセス起動を実現できない。
	StartProgramAndActivateFile( sFilePath, bShowToolTip:=True ) ; {{{
	{
		;*** preprocess ***
		If ( sFilePath == "" )
		{
			MsgBox "[ERROR] please specify arguments to StartProgramAndActivateFile()."
			return
		}
		sFileName := ExtractFileName(sFilePath)
		;MsgBox "sFilePath = " . sFilePath . "`nsFileName = " . sFileName
		
		;*** show tooltip ***
		If ( bShowToolTip == True ) {
			ShowAutoHideToolTip(sFileName . " is starting...", giSTART_PRG_TOOLTIP_SHOW_TIME_MS)
		}
		
		;*** start program ***
		Try {
			Run sFilePath, , , &sOutputVarPID
			WinWait("ahk_pid " . sOutputVarPID, , 1)
			WinActivate "ahk_pid " . sOutputVarPID
		;	BringActiveWindowToTop()
		} Catch Error as err {
		;	MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
		;		, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
		}
		ToolTip()
		return
	} ; }}}
	
	; 起動＆アクティベート処理 (実行プログラム指定のみ)
	StartProgramAndActivateExe( sExePath, bLaunchSingleProcess:=False, bShowToolTip:=True ) ; {{{
	{
		;*** preprocess ***
		If ( sExePath == "" )
		{
			MsgBox "[ERROR] please specify arguments to StartProgramAndActivateExe()."
			return
		}
		
		sExeName := ExtractFileName(sExePath)
		sExeDirPath := ExtractDirPath(sExePath)
		;MsgBox "sExePath = " . sExePath . "`nsExeName = " . sExeName . "`nsExeDirPath = " . sExeDirPath . "`nbLaunchSingleProcess = " . bLaunchSingleProcess
		
		;*** show tooltip ***
		If ( bShowToolTip == True ) {
			ShowAutoHideToolTip(sExeName . " is starting...", giSTART_PRG_TOOLTIP_SHOW_TIME_MS)
		}
		
		;*** check if the program is running ***
		If ( bLaunchSingleProcess == True ) {
			iPID := ProcessExist(sExeName)
			If (iPID != 0)
			{
				WinActivate "ahk_pid " . iPID
				return
			}
		}
		
		;*** start program ***
		Try {
			Run sExePath, sExeDirPath, , &sOutputVarPID
			WinWait("ahk_pid " . sOutputVarPID, , 1)
			WinActivate "ahk_pid " . sOutputVarPID
		;	BringActiveWindowToTop()
		} Catch Error as err {
		;	MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
		;		, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
		}
		ToolTip()
		return
	} ; }}}
	BringActiveWindowToTop() ; {{{
	{
		Try {
			;常に最前面ON→OFFにより、アクティブウィンドウを最前面に設定する
			Sleep 100
			WinSetAlwaysOnTop 1, "A" ; 常に最前面ON
			Sleep 100
			WinSetAlwaysOnTop 0, "A" ; 常に最前面OFF
		} Catch Error as err {
		;	MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
		;		, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
		}
	} ; }}}

	; 今押している修飾キーと共にキー送信する
	SendKeyWithModKeyCurPressing( sSendKey ) ; {{{
	{
		bIsPressShift := GetKeyState("Shift","P")
		bIsPressCtrl := GetKeyState("Ctrl","P")
		bIsPressAlt := GetKeyState("Alt","P")
		if(bIsPressShift and bIsPressCtrl and bIsPressAlt){
			Send "!^+{" . sSendKey . "}"
		} else if(bIsPressShift and bIsPressCtrl){
			Send "^+{" . sSendKey . "}"
		} else if(bIsPressShift and bIsPressAlt){
			Send "!+{" . sSendKey . "}"
		} else if(bIsPressAlt and bIsPressCtrl){
			Send "!^{" . sSendKey . "}"
		} else if(bIsPressAlt){
			Send "!{" . sSendKey . "}"
		} else if(bIsPressCtrl){
			Send "^{" . sSendKey . "}"
		} else if(bIsPressShift){
			Send "+{" . sSendKey . "}"
		} else {
			Send "{" . sSendKey . "}"
		}
		return
	} ; }}}

	;Windowタイル切り替え
	InitWinTileMode() ; {{{
	{
		ClearWinTileMode()
		SetTimerClearWinTileMode()
	} ; }}}
	IncrementWinTileMode() ; {{{
	{
		global giWinTileMode
		giWinTileMode += 1
		iWinTileModeMin := GetWinTileModeMin()
		iWinTileModeMax := GetWinTileModeMax()
		if ( giWinTileMode > iWinTileModeMax ) {
			giWinTileMode := iWinTileModeMin
		} else {
			giWinTileMode := CropValue(giWinTileMode, iWinTileModeMin, iWinTileModeMax)
		}
	;	MsgBox "[DBG] IncrementWinTileMode()" . "`ngiWinTileMode = " . giWinTileMode
	} ; }}}
	DecrementWinTileMode() ; {{{
	{
		global giWinTileMode
		giWinTileMode -= 1
		iWinTileModeMin := GetWinTileModeMin()
		iWinTileModeMax := GetWinTileModeMax()
		if ( giWinTileMode < iWinTileModeMin ) {
			giWinTileMode := iWinTileModeMax
		} else {
			giWinTileMode := CropValue(giWinTileMode, iWinTileModeMin, iWinTileModeMax)
		}
	;	MsgBox "[DBG] DecrementWinTileMode()" . "`ngiWinTileMode = " . giWinTileMode
	} ; }}}
	GetWinTileModeMin() ; {{{
	{
		return CropWinTileModeWithMonNum(giWIN_TILE_MODE_RANGE_MIN)
	} ; }}}
	GetWinTileModeMax() ; {{{
	{
		return CropWinTileModeWithMonNum(giWIN_TILE_MODE_RANGE_MAX)
	} ; }}}
	CropWinTileModeWithMonNum(iInWinTileMode) ; {{{
	{
		iOutWinTileMode := giWIN_TILE_MODE_INIT
		iMonitorNum := GetMonitorNum()
		switch iMonitorNum
		{
			case 1:		iOutWinTileMode := CropValue(iInWinTileMode, 4, 6) ; Main only
			case 2:		iOutWinTileMode := CropValue(iInWinTileMode, 1, 6) ; Main + Sub
			default:	MsgBox "[error] invalid iMonitorNum : " . iMonitorNum
		}
		return iOutWinTileMode
	} ; }}}
	SetTimerClearWinTileMode() ; {{{
	{
		SetTimer ClearWinTileMode, giWIN_TILE_MODE_CLEAR_INTERVAL_MS
	} ; }}}
	ClearWinTileMode() ; {{{
	{
		global giWinTileMode
		giWinTileMode := giWIN_TILE_MODE_INIT
		;ShowAutoHideTrayTip("タイルモードクリアタイマー", "タイルモードをクリアしました", 5000)
		Return
	} ; }}}
	; ウィンドウサイズ切り替え
	ApplyWinTileMode() ; {{{
	{
		global giWinTileMode
		GetMonitorPosInfo(1, &dX1, &dY1, &dWidth1, &dHeight1 )
		GetMonitorPosInfo(2, &dX2, &dY2, &dWidth2, &dHeight2, "Bottom", giWIN_TILE_MODE_WIN_RANGE_RATE )
	;	MsgBox "[DBG] ApplyWinTileMode() " .
	;		"`n giWinTileMode = " . giWinTileMode .
	;		"`n dX1 = " . dX1 . "`n dY1 = " . dY1 . "`n dWidth1 = " . dWidth1 . "`n dHeight1 = " . dHeight1 .
	;		"`n dX2 = " . dX2 . "`n dY2 = " . dY2 . "`n dWidth2 = " . dWidth2 . "`n dHeight2 = " . dHeight2
		
		switch giWinTileMode
		{
			case 1:		MoveActiveWin(dX2, dY2, dWidth2, dHeight2)
			case 2:		MoveActiveWin(dX2, dY2, dWidth2, dHeight2, "Top")
			case 3:		MoveActiveWin(dX2, dY2, dWidth2, dHeight2, "Bottom")
			case 4:		MoveActiveWin(dX1, dY1, dWidth1, dHeight1)
			case 5:		MoveActiveWin(dX1, dY1, dWidth1, dHeight1, "Left")
			case 6:		MoveActiveWin(dX1, dY1, dWidth1, dHeight1, "Right")
			default:	MsgBox "[error] invalid giWinTileMode : " . giWinTileMode
		}
		return
	} ; }}}
	GetMonitorNum() ; {{{
	{
		return SysGet(80) ; SM_CMONITORS: Number of display monitors on the desktop (not including "non-display pseudo-monitors").
	} ; }}}
	GetMonitorPosInfo( iMonIdx, &dX, &dY, &dWidth, &dHeight, sAttachSide:="", iWinRangeRate:=0 ) ; {{{
	{
		iMonNum := GetMonitorNum()
		if ( iMonIdx > iMonNum)
		{
			return False
		}
		
		try
		{
			ActualN := MonitorGetWorkArea(iMonIdx, &Left, &Top, &Right, &Bottom)
		;	MsgBox "Left: " Left " -- Top: " Top " -- Right: " Right " -- Bottom: " Bottom
		} Catch Error as err {
			MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
				, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
		}
		dY := Top
		if ( Left < Right ) {
			dX := Left
			dWidth := Right - Left + 1
		} else {
			dX := Right
			dWidth := Left - Right + 1
		}
		dHeight := Bottom - Top + 1
	;	MsgBox "[DBG] GetMonitorPosInfo() 01" . "`n iMonIdx = " . iMonIdx . "`n dX = " . dX . "`n dY = " . dY . "`n dWidth = " . dWidth . "`n dHeight = " . dHeight
		
		switch sAttachSide
		{
			case "Top":
				dHeight := dHeight * iWinRangeRate
			case "Bottom":
				dY := dY + ( dHeight * ( 1 - iWinRangeRate) )
				dHeight := dHeight * iWinRangeRate
			case "Left":
				dWidth := dWidth * iWinRangeRate
			case "Right":
				dX := dX + ( dWidth * ( 1 - iWinRangeRate) )
				dWidth := dWidth * iWinRangeRate
			default:
				; Do Nothing
		}
	;	MsgBox "[DBG] GetMonitorPosInfo() 02" . "`n iMonIdx = " . iMonIdx . "`n dX = " . dX . "`n dY = " . dY . "`n dWidth = " . dWidth . "`n dHeight = " . dHeight
		return True
	} ; }}}
	MoveActiveWin(iInX, iInY, iInWidth, iInHeight, sOutputSide:="") ; {{{
	{
		switch sOutputSide
		{
			case "Top":
				iWinX		:= Integer(iInX)
				iWinY		:= Integer(iInY)
				iWinWidth	:= Integer(iInWidth)
				iWinHeight	:= Integer(iInHeight / 2)
			case "Bottom":
				iWinX		:= Integer(iInX)
				iWinY		:= Integer(iInY + iInHeight / 2)
				iWinWidth	:= Integer(iInWidth)
				iWinHeight	:= Integer(iInHeight / 2)
			case "Left":
				iWinX		:= Integer(iInX)
				iWinY		:= Integer(iInY)
				iWinWidth	:= Integer(iInWidth / 2)
				iWinHeight	:= Integer(iInHeight)
			case "Right":
				iWinX		:= Integer(iInX + (iInWidth / 2))
				iWinY		:= Integer(iInY)
				iWinWidth	:= Integer(iInWidth / 2)
				iWinHeight	:= Integer(iInHeight)
			default:
				iWinX		:= Integer(iInX)
				iWinY		:= Integer(iInY)
				iWinWidth	:= Integer(iInWidth)
				iWinHeight	:= Integer(iInHeight)
		}
	;	MsgBox "[DBG] MoveActiveWin() " .
	;		"`n iWinX = " . iWinX . 
	;		"`n iWinY = " . iWinY . 
	;		"`n iWinWidth = " . iWinWidth . 
	;		"`n iWinHeight = " . iWinHeight . 
	;		""
		Try {
			WinMove iWinX, iWinY, iWinWidth, iWinHeight, "A"
		} Catch Error as err {
		;	MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
		;		, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
		}
	} ; }}}

	; ファイル名取得
	ExtractFileName( sFilePath ) ; {{{
	{
		SplitPath sFilePath, &sFileName, &sDirPath, &sExtName, &sFileBaseName, &sDrive
		sFileName := StrReplace(sFileName, "`"", )
	;	MsgBox sFilePath . "`n" . sFileName . "`n" . sDirPath . "`n" . sExtName . "`n" . sFileBaseName . "`n" . sDrive
		return sFileName
	} ; }}}
	; ディレクトリパス取得
	ExtractDirPath( sFilePath ) ; {{{
	{
		SplitPath sFilePath, &sFileName, &sDirPath, &sExtName, &sFileBaseName, &sDrive
		sDirPath := StrReplace(sDirPath, "`"", )
	;	MsgBox sFilePath . "`n" . sFileName . "`n" . sDirPath . "`n" . sExtName . "`n" . sFileBaseName . "`n" . sDrive
		return sDirPath
	} ; }}}

	; 選択ファイルパス取得＠explorer
	GetSelFilePathAtExplorer( bIsDelimiterSpace ) ; {{{
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
	;	MsgBox "sTrgtPaths = " . sTrgtPaths
		return sTrgtPaths
	} ; }}}
	; 現在フォルダパス取得＠explorer
	GetCurDirPathAtExplorer() ; {{{
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
	;	MsgBox "sTrgtPaths = " . sTrgtPaths
		return sTrgtPaths
	} ; }}}
	; 選択ファイル名取得＠explorer
	GetSelFileNameAtExplorer() ; {{{
	{
		sFilePaths := GetSelFilePathAtExplorer(0)
		sDirPaths := GetCurDirPathAtExplorer()
		sTrgtPaths := StrReplace(sFilePaths, sDirPaths . "\", )
		sTrgtPaths := StrReplace(sTrgtPaths, "`"", )
	;	MsgBox "sTrgtPaths = " . sTrgtPaths . "`nsFilePaths = " . sFilePaths . "`nsDirPaths = " . sDirPaths
		return sTrgtPaths
	} ; }}}
	; ファイルリストへフォーカスを移す＠explorer
	;FocusFileDirListAtExplorer() ; {{{
	;{
	;	Sleep 100
	;	ControlFocus, SysTreeView321
	;	If ErrorLevel=0
	;	{
	;	;	MsgBox "ControlFocus success"
	;		Sleep 100
	;		Send "{Tab}"
	;	}
	;} ; }}}
	FocusFileDirListAtExplorer() ; {{{
	{
		WinActivate "ahk_class CabinetWClass ahk_exe Explorer.EXE"
		Sleep 200
		Send "^f"
		Sleep 200
		Send "{Tab}"
		Sleep 200
		Send "{Tab}"
		return
	} ; }}}

	; ツールチップ表示
	; （カーソル付近に表示されるメッセージ）
	ShowAutoHideToolTip(sMsg, iShowPeriodMs) ; {{{
	{
		ToolTip(sMsg)
		SetTimer () => ToolTip(), -1 * iShowPeriodMs
		Return
	} ; }}}

	; トレイチップ表示
	; （Windowsのタスクトレイ付近に表示されるメッセージ）
	ShowAutoHideTrayTip(sTitle, sMsg, iShowPeriodMs) ; {{{
	{
		TrayTip sMsg, sTitle, 1
		SetTimer () => TrayTip(), -1 * iShowPeriodMs
		Return
	} ; }}}

	; 画面明るさ設定
	InitScreenBrightness() ; {{{
	{
		global giBrightness
		iNowHour := Integer(FormatTime(A_Now, "H"))
		if (giSCREEN_BRIGHTNESS_DAY_START_TIME < iNowHour && iNowHour < giSCREEN_BRIGHTNESS_DAY_END_TIME)
		{
			giBrightness := giSCREEN_BRIGHTNESS_INIT_DAY
		} else {
			giBrightness := giSCREEN_BRIGHTNESS_INIT_NIGHT
		}
		global gasDimId := Array()
		iMonitorCount := MonitorGetCount()
	;	MsgBox "iMonitorCount = " . iMonitorCount . ", giBrightness = " . giBrightness
		Loop iMonitorCount
		{
			MonitorGet(A_Index, &MonitorLeft, &MonitorTop, &MonitorRight, &MonitorBottom)
			Width := MonitorRight - MonitorLeft
			Height := MonitorBottom - MonitorTop
			oDimGui := Gui()
			oDimGui.Opt("+LastFound +ToolWindow -Disabled -SysMenu -Caption +E0x20 +AlwaysOnTop")
			oDimGui.BackColor := "000000"	;フィルタの色（HTMLカラーコード参照）
			oDimGui.Title := "DimMonitor" . A_Index
			oDimGui.Show("X" . MonitorLeft . " Y" . MonitorTop . " W" . Width . " H" . Height)
			gasDimId.push WinGetId("DimMonitor" . A_Index . " ahk_class AutoHotkeyGUI")
			iDimId := gasDimId[A_Index]
			iTransparency := 100 - giBrightness
			WinSetTransparent(Integer(iTransparency * 255 / 100), "ahk_id " . iDimId)
		;	MsgBox "iMonitorCount = " . iMonitorCount . ", A_Index = " . A_Index . ", iDimId = " . iDimId
		}
		Return
	} ; }}}
	BrightenScreen() ; {{{
	{
		global giBrightness
		giBrightness += giSCREEN_BRIGHTNESS_STEP
		if (giBrightness > giSCREEN_BRIGHTNESS_MAX)
		{
			giBrightness := giSCREEN_BRIGHTNESS_MAX
		}
		ApplyBrightness(True)
	} ; }}}
	DarkenScreen() ; {{{
	{
		global giBrightness
		giBrightness -= giSCREEN_BRIGHTNESS_STEP
		if (giBrightness < giSCREEN_BRIGHTNESS_MIN)
		{
			giBrightness := giSCREEN_BRIGHTNESS_MIN
		}
		ApplyBrightness(True)
	} ; }}}
	FlashScreen(iDarkBrightness:=10, iFlashCount:=10, iFlashIntervalMs:=50, iFlashSleepMs:=200) ; {{{
	{
		iFlashIntervalMs := 50
		iFlashSleepMs := 100
		Loop iFlashCount
		{
			SetBrightnessTemporary(iDarkBrightness, iFlashIntervalMs)
			sleep iFlashSleepMs
		}
	} ; }}}
	SetBrightness(iBrightness) ; {{{
	{
		global giBrightness
		giBrightness := iBrightness
		ApplyBrightness(True)
	} ; }}}
	SetBrightnessTemporary(iBrightness, iWaitTimeMs) ; {{{
	{
		global giBrightness
		global giBrightnessOld := giBrightness
		giBrightness := iBrightness
		ApplyBrightness(False)
		SetTimer(SetOldBrightness, -1 * iWaitTimeMs)
	} ; }}}
	SetOldBrightness() ; {{{
	{
		global giBrightness
		global giBrightnessOld
		giBrightness := giBrightnessOld
		ApplyBrightness(False)
	} ; }}}
	ApplyBrightness(bShowToolTip:=True) ; {{{
	{
		global giBrightness
		iMonitorCount := MonitorGetCount()
		Loop iMonitorCount
		{
			iDimId := gasDimId[A_Index]
			iTransparency := 100 - giBrightness
			WinSetTransparent(Integer(iTransparency * 255 / 100), "ahk_id " . iDimId)
		}
		If ( bShowToolTip == True ) {
			ShowAutoHideToolTip("明るさ：" . giBrightness . "%", 500)
		}
		Return
	} ; }}}

	;クリップボード設定
	SetClipboard(sStr) ; {{{
	{
		A_Clipboard := ""
		A_Clipboard := sStr
		ClipWait
	} ; }}}

	; GUI
	CreateSlctCmndWindowZip() ; {{{
	{
		global gmyGui
		global gogcListBoxAnswer
		gmyGui := Gui()
		gmyGui.Add("Text", , "圧縮/パスワード圧縮/解凍を実行します。`n処理を選択してください。")
		gogcListBoxAnswer := gmyGui.Add("ListBox", "vAnswer Choose1 R3", ["圧縮", "パスワード圧縮", "解凍"])
		ogcButtonZipEnter := gmyGui.Add("Button", "Hidden w0 h0 Default", "ZipEnter")
		ogcButtonZipEnter.OnEvent("Click", EventClickAtZip.Bind("Normal"))
		gmyGui.OnEvent("Close", EventEscape)
		gmyGui.OnEvent("Escape", EventEscape)
		gmyGui.Show("Center")
	} ; }}}
	EventClickAtZip(A_GuiEvent, GuiCtrlObj, Info, *) ; {{{
	{
		global gmyGui
		global gogcListBoxAnswer
		vAnswer := gogcListBoxAnswer.Text
		;MsgBox vAnswer
		gmyGui.Destroy()
		sTrgtPaths := GetSelFilePathAtExplorer(1)
		sDirPath := EnvGet("MYDIRPATH_CODES")
		switch vAnswer
		{
			case "圧縮":
				RunWait(sDirPath . "\vbs\tools\7zip\ZipFile.vbs " . sTrgtPaths)
			case "パスワード圧縮":
				RunWait(sDirPath . "\vbs\tools\7zip\ZipPasswordFile.vbs " . sTrgtPaths)
			case "解凍":
				RunWait(sDirPath . "\vbs\tools\7zip\UnzipFile.vbs " . sTrgtPaths)
			default:
				MsgBox "[ERROR] 圧縮/パスワード圧縮/解凍 選択"
		}
		;FocusFileDirListAtExplorer()
		return
	} ; }}}
	CreateSlctCmndWindowLink() ; {{{
	{
		global gmyGui
		global gogcListBoxAnswer
		gmyGui := Gui()
		gmyGui.Add("Text", , "ショートカット/シンボリックリンクを作成します。`n処理を選択してください。")
		gogcListBoxAnswer := gmyGui.Add("ListBox", "vAnswer Choose1 R2", ["ショートカット作成", "シンボリックリンク作成"])
		ogcButtonSelLinkEnter := gmyGui.Add("Button", "Hidden w0 h0 Default", "SelLinkEnter")
		ogcButtonSelLinkEnter.OnEvent("Click", EventClickAtLink.Bind("Normal"))
		gmyGui.OnEvent("Close", EventEscape)
		gmyGui.OnEvent("Escape", EventEscape)
		gmyGui.Show("Center")
	} ; }}}
	EventClickAtLink(A_GuiEvent, GuiCtrlObj, Info, *) ; {{{
	{
		global gmyGui
		global gogcListBoxAnswer
		vAnswer := gogcListBoxAnswer.Text
		gmyGui.Destroy()
		sTrgtPaths := GetSelFilePathAtExplorer(1)
		sDirPath := EnvGet("MYDIRPATH_CODES")
		switch vAnswer
		{
			case "ショートカット作成":
				RunWait(sDirPath . "\vbs\command\CreateShortcutFile.vbs " . sTrgtPaths . ".lnk " . sTrgtPaths)
			case "シンボリックリンク作成":
				RunWait(sDirPath . "\vbs\tools\win\file_ope\CreateSymbolicLink.vbs " . sTrgtPaths)
			default:
				MsgBox "[ERROR] ショートカット/シンボリックリンク作成"
		}
		;FocusFileDirListAtExplorer()
		return
	} ; }}}
	CreateSlctCmndWindowPathList() ; {{{
	{
		global gmyGui
		global gogcListBoxAnswer
		gmyGui := Gui()
		gmyGui.Add("Text", , "パス一覧を作成します。`n処理を選択してください。")
		gogcListBoxAnswer := gmyGui.Add("ListBox", "vAnswer Choose1 R4", ["ファイル＆フォルダ一覧作成", "ファイル一覧作成", "フォルダ一覧作成", "フォルダツリー作成"])
		ogcButtonPathListEnter := gmyGui.Add("Button", "Hidden w0 h0 Default", "PathListEnter")
		ogcButtonPathListEnter.OnEvent("Click", EventClickPathList.Bind("Normal"))
		gmyGui.OnEvent("Close", EventEscape)
		gmyGui.OnEvent("Escape", EventEscape)
		gmyGui.Show("Center")
	} ; }}}
	EventClickPathList(A_GuiEvent, GuiCtrlObj, Info, *) ; {{{
	{
		global gmyGui
		global gogcListBoxAnswer
		vAnswer := gogcListBoxAnswer.Text
		gmyGui.Destroy()
		sDirPath := GetCurDirPathAtExplorer()
		switch vAnswer
		{
			case "ファイル＆フォルダ一覧作成":
				RunWait(A_ComSpec " /c dir /s /b /a > `"" sDirPath "\_PathList_FileDir.txt`"", sDirPath)
			case "ファイル一覧作成":
				RunWait(A_ComSpec " /c dir *.* /b /s /a:a-d > `"" sDirPath "\_PathList_File.txt`"", sDirPath)
			case "フォルダ一覧作成":
				RunWait(A_ComSpec " /c dir /b /s /a:d > `"" sDirPath "\_PathList_Dir.txt`"", sDirPath)
			case "フォルダツリー作成":
				RunWait(A_ComSpec " /c tree /f > `"" sDirPath "\_DirTree.txt`"", sDirPath)
			default:
				MsgBox "[ERROR] パス一覧作成"
		}
		;FocusFileDirListAtExplorer()
	} ; }}}
	EventEscape(*) ; {{{
	{
		global gmyGui
	;	MsgBox "エスケープされました"
		gmyGui.Destroy()
	} ; }}}

	; IME.ahk
	; [URL] https://github.com/s-show/AutoHotKey/blob/AutoHotKey/IME.ahk
	;-----------------------------------------------------------
	; IMEの状態の取得
	;   WinTitle="A"    対象Window
	;   戻り値          1:ON / 0:OFF
	;-----------------------------------------------------------
	IME_GET(WinTitle:="A")  { ; {{{
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
	} ; }}}
	;-----------------------------------------------------------
	; IMEの状態をセット
	;	SetSts			1:ON / 0:OFF
	;	WinTitle="A"	対象Window
	;	戻り値			0:成功 / 0以外:失敗
	;-----------------------------------------------------------
	IME_SET(SetSts, WinTitle:="A")    { ; {{{
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
	} ; }}}

	; 本スクリプトをリロードする
	ReloadMe() ; {{{
	{
		Reload
		Sleep 1000 ; リロードに成功した場合、リロードはスリープ中にこのインスタンスを閉じるので、以下の行に到達することはない
		MsgBox "スクリプト" . A_ScriptName . "の再読み込みに失敗しました"
	} ; }}}

	; 今月/先月の月日を取得する
	StoreCurYearMonths() ; {{{
	{
		global gsYearCur := ""
		global gsMonth1DegCur := ""
		global gsMonth2DegCur := ""
		global gsYearLast := ""
		global gsMonth1DegLast := ""
		global gsMonth2DegLast := ""
		GetYearMonth(A_YYYY, A_MM, &gsYearCur, &gsMonth1DegCur, &gsMonth2DegCur)
		GetYearMonth(A_YYYY, A_MM, &gsYearLast, &gsMonth1DegLast, &gsMonth2DegLast, -1)
		;MsgBox gsYearCur . "/" . gsMonth1DegCur . "," . gsMonth2DegCur . "`n" . gsYearLast . "/" . gsMonth1DegLast . "," . gsMonth2DegLast
	} ; }}}
	; 月日を取得する
	GetYearMonth(sInYear, sInMonth, &sOutYear, &sOutMonth1Deg, &sOutMonth2Deg, iOffset:=0 ) ; {{{
	{
		if (iOffset > 12 || iOffset < -12)
		{
			MsgBox "[error] GetYearMonth() iOffset is " . iOffset . ". iOffset must be keep within -12~12."
			return
		}
		;MsgBox sInYear . "/" . sInMonth
		
		iTrgtMonth := Integer(sInMonth) + iOffset
		iTrgtYear := Integer(sInYear)
		if (iTrgtMonth < 1)
		{
			iTrgtMonth := 12 + iTrgtMonth
			iTrgtYear := iTrgtYear - 1
		}
		else if (iTrgtMonth > 12)
		{
			iTrgtMonth := iTrgtMonth - 12
			iTrgtYear := iTrgtYear + 1
		}
		else
		{
			; Do Nothing
		}
		;MsgBox String(iTrgtYear) . "/" . String(iTrgtMonth)
		sOutYear := Format("{1:04d}" , String(iTrgtYear))
		sOutMonth1Deg := Format("{1:d}" , String(iTrgtMonth))
		sOutMonth2Deg := Format("{1:02d}" , String(iTrgtMonth))
	} ; }}}
		Test_GetYearMonth() { ; {{{
			sYear := ""
			sMonth1Deg := ""
			sMonth2Deg := ""
			sOutStr := ""
			sInYear := "2022"
			sInMonth := "01"
			iTestCase := 1
			
			if (iTestCase == 0) {
				; normal case
				GetYearMonth(A_YYYY, A_MM, &sYear, &sMonth1Deg, &sMonth2Deg)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, 0)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, 1)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, 2)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, 12)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, -1)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, -2)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, -12)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
			} else {
				; error case
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, 13)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
				
				GetYearMonth(sInYear, sInMonth, &sYear, &sMonth1Deg, &sMonth2Deg, -13)
				sOutStr := sOutStr . "`n" . sYear . "/" . sMonth1Deg . "," . sMonth2Deg
			}
			
			MsgBox sOutStr
		} ; }}}

	; Window最前面化
	ToggleAlwaysOnTopEnable() ; {{{
	{
		static bEnableAlwaysOnTop := False
		WinSetAlwaysOnTop -1, "A"
		sActiveWinTitle := WinGetTitle("A")
		if (bEnableAlwaysOnTop == False)
		{
			ShowAutoHideTrayTip("", "Window最前面を【有効】にします`n`n" . sActiveWinTitle, 2000)
			bEnableAlwaysOnTop := True
		}
		else
		{
			ShowAutoHideTrayTip("", "Window最前面を【解除】します`n`n" . sActiveWinTitle, 2000)
			bEnableAlwaysOnTop := False
		}
	} ; }}}

	; ウィンドウスリープ抑制
	InitSleepPreventing() ; {{{
	{
		SetSleepPreventingMode("Disable", False)
	} ; }}}
	SetSleepPreventingMode(sMode, bShowToolTip:=True) ; {{{
	{
		static bEnablePreventWindow
		switch sMode {
			case "Toggle":
				if (bEnablePreventWindow = False) {
					bEnablePreventWindow := True
				} else {
					bEnablePreventWindow := False
				}
			case "Enable":
				bEnablePreventWindow := True
			case "Disable":
				bEnablePreventWindow := False
			default:
				MsgBox "[ERROR] SetSleepPreventing() unknown mode : " . sMode
		}
		
		if (bEnablePreventWindow = True) {
			if (bShowToolTip == True) {
				ShowAutoHideTrayTip("", giSLEEPPREVENT_PROGRAM_NAME . " のスリープ抑制を【有効化】します", 2000)
			}
			SetTimer ActivateTargetWindow, giSLEEPPREVENT_INTERVAL_TIME_MS
		} else {
			if (bShowToolTip == True) {
				ShowAutoHideTrayTip("", giSLEEPPREVENT_PROGRAM_NAME . " のスリープ抑制を【解除】します", 2000)
			}
			SetTimer ActivateTargetWindow, 0
		}
	} ; }}}
	ActivateTargetWindow() ; {{{
	{
		if (gbSLEEPPREVENT_SHOW_TRAYTIP_WITH_ACT == True) {
			ShowAutoHideTrayTip("", giSLEEPPREVENT_PROGRAM_NAME . " アクティベート実行", 2000)
		}
		Try {
			iActiveWindowIdOld := WinGetID("A")
			WinActivate "ahk_exe " . giSLEEPPREVENT_EXE_NAME
			Send giSLEEPPREVENT_KEY_NAME
			Sleep 200
			WinActivate "ahk_id " . iActiveWindowIdOld
		} Catch Error as err {
			;MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
			;	, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
		}
	} ; }}}

	;モニタ中心にカーソル移動
	MoveCursolToMonitorCenter() { ; {{{
		static iMoveTrgtMonNum := 1
		GetMonitorPosInfo( iMoveTrgtMonNum, &dX, &dY, &dWidth, &dHeight )
		dCurX := dX + Integer(dWidth / 2)
		dCurY := dY + Integer(dHeight / 2)
		CoordMode "Mouse", "Screen"
		MouseMove dCurX, dCurY
		CoordMode "Mouse"
		Sleep 100
		Send("{Ctrl down}")
		Sleep 100
		Send("{Ctrl up}")
		
		iMonNum := GetMonitorNum()
		If ( iMoveTrgtMonNum >= iMonNum ) {
			iMoveTrgtMonNum := 1
		} Else {
			iMoveTrgtMonNum := iMoveTrgtMonNum + 1
		}
	} ; }}}

	; カーソル移動
	MoveCursor(sDirection) ; {{{
	{
		iMoveOffset := 0
		if (GetKeyState("Shift","P")) {
			iMoveOffset := giMOVECURSOL_MOVE_OFFSET_FAR
		} else {
			iMoveOffset := giMOVECURSOL_MOVE_OFFSET_NEAR
		}
		switch sDirection {
			case "Left":
				MouseMove -iMoveOffset, 0, 0, "R"
			case "Right":
				MouseMove iMoveOffset, 0, 0, "R"
			case "Up":
				MouseMove 0, -iMoveOffset, 0, "R"
			case "Down":
				MouseMove 0, iMoveOffset, 0, "R"
			default:
				MsgBox "[ERROR] MoveCursor() unknown direction : " . sDirection
		}
	} ; }}}

	; 右Alt->AppsKey置換え
	InitRAltAppsKeyMode() { ; {{{
		global gbReplaceRAlt2AppsKey
		gbReplaceRAlt2AppsKey := gbRALT2APPSKEY_RALT_TO_APPSKEY
	} ; }}}
	SwitchRAltAppsKeyMode() { ; {{{
		global gbReplaceRAlt2AppsKey
		if (gbReplaceRAlt2AppsKey) {
			ShowAutoHideTrayTip("", "右Alt->AppsKey置換えを【無効】にします。", 2000)
			gbReplaceRAlt2AppsKey := false
		} else {
			ShowAutoHideTrayTip("", "右Alt->AppsKey置換えを【有効】にします。", 2000)
			gbReplaceRAlt2AppsKey := true
		}
	} ; }}}
	PressRAlt() ; {{{
	{
		global gbReplaceRAlt2AppsKey
		if (gbReplaceRAlt2AppsKey) {
			Send "{AppsKey}"	;右Altキーをコンテキストメニュー表示に変更
		} else {
			Send "{RAlt}"
		}
	} ; }}}

	; キッチンタイマー
	ClearKitchenTimer() ; {{{
	{
		sLogFilePath := EnvGet("MYDIRPATH_DESKTOP") . "\KitchenTimer_*.log"
		DeleteFile(sLogFilePath)
	} ; }}}
	SetKitchenTimer(fIntervalMin:=0.0, bShowMsgs:=true, bCreateLogFile:=true) ; {{{
	{
		kitchen_timer := KitchenTimer(bShowMsgs, bCreateLogFile)
		kitchen_timer.SetTimeWithIntervalMin(fIntervalMin)
		kitchen_timer.Start()
	} ; }}}
	RestartKitchenTimer() ; {{{
	{
		sFilePattern := EnvGet("MYDIRPATH_DESKTOP") . "\KitchenTimer_*.log"
		Loop Files, sFilePattern
		{
			sFileName := A_LoopFileName
			sFilePath := EnvGet("MYDIRPATH_DESKTOP") . "\" . sFileName
			
			kitchen_timer := KitchenTimer(false, true)
			kitchen_timer.SetTimeWithLogFile(sFilePath)
			kitchen_timer.Start()
		}
		
	} ; }}}
	class KitchenTimer { ; {{{
		__New(bShowMsgs:=true, bCreateLogFile:=true) { ; {{{
			this.bShowMsgs := bShowMsgs
			this.bCreateLogFile := bCreateLogFile
			this.objCallbackFunc := ObjBindMethod(this, "TimerCallback")
			this.fIntervalMin := 0.0
			this.sLogFilePath := ""
			this.bIsRestart := false
			this.bReadyToStart := false
		} ; }}}
		SetTimeWithIntervalMin(fIntervalMin:=0.0) { ; {{{
			this.bIsRestart := false
			this.fOrigIntervalMin := 0.0
			this.sOrigStartDateTime := ""
			
			; 時間設定
			if (fIntervalMin = 0.0) {
				; 初期値取得
				if (gbKITCHENTIMER_SAVE_INIT_MIN) {
					sConfigDirPath := EnvGet("MYDIRPATH_DOCUMENTS") . "\" . gsCONFIG_DIR_NAME
					sConfigFilePath := sConfigDirPath . "\" . gsKITCHENTIMER_CONFIG_FILE_NAME
					if (FileExist(sConfigFilePath)) {
						sFileLines := FileRead(sConfigFilePath)
						asFileLines := StrSplit(sFileLines, "`n")
						fIntervalMin := Float(asFileLines[1])
					} else {
						fIntervalMin := gfKITCHENTIMER_INIT_MIN
					}
				} else {
					fIntervalMin := gfKITCHENTIMER_INIT_MIN
				}
				
				; 時間設定
				Try {
					InputBoxObj := InputBox("キッチンタイマーを設定します。`n時間[分]を設定してください。", "キッチンタイマー", , String(fIntervalMin))
					if (InputBoxObj.Result = "Cancel") {
						MsgBox "キャンセルされたため、処理を中断します。"
						Return
					}
					sIntervalMin := InputBoxObj.Value
					fIntervalMin := Float(sIntervalMin)
					if (fIntervalMin <= 0.0) {
						throw
					}
				} Catch Error as err {
					MsgBox "不正な時間が指定されました。`n" . sIntervalMin . "`n`n処理を中断します。"
					Return
				}
				
				; 初期値保存
				if (gbKITCHENTIMER_SAVE_INIT_MIN) {
					DirCreate sConfigDirPath
					sFileContents := String(fIntervalMin)
					DeleteFile(sConfigFilePath)
					FileAppend sFileContents, sConfigFilePath
				}
			}
			If (fIntervalMin <= 0.0) {
				MsgBox "不正な時間が指定されました。`n" . fIntervalMin . "`n`n処理を中断します。"
				Return
			}
			this.fIntervalMin := fIntervalMin
			this.bReadyToStart := true
		} ; }}}
		SetTimeWithLogFile(sLogFilePath) { ; {{{
			this.bIsRestart := true
			if (!FileExist(sLogFilePath)) {
				MsgBox "ファイルが存在しません。`n  " . sLogFilePath . "`n`n処理を中断します。"
				Return
			}
			sFileLines := FileRead(sLogFilePath)
			asFileLines := StrSplit(sFileLines, "`n")
			sTargetDateTime := asFileLines[1]
			fOrigIntervalMin := Float(asFileLines[2])
			sOrigStartDateTime := asFileLines[3]
			;MsgBox sTargetDateTime . "`n" . String(fOrigIntervalMin) . "`n" . sOrigStartDateTime
			
			this.fOrigIntervalMin := fOrigIntervalMin
			this.sOrigStartDateTime := sOrigStartDateTime
			
			DeleteFile(sLogFilePath)
			
			fIntervalMin := 0.0
			sStartDateTime := A_Now
			iElapsedSecond := DateDiff(sTargetDateTime, sStartDateTime, "Seconds")
			;MsgBox sTargetDateTime . "`n" . sStartDateTime . "`n" . iElapsedSecond
			if (iElapsedSecond > 0) {
				fIntervalMin := Float(iElapsedSecond / 60)
			}
			;MsgBox fIntervalMin
			this.fIntervalMin := fIntervalMin
			this.bReadyToStart := true
		} ; }}}
		Start() { ; {{{
			; 事前チェック
			If (!this.bReadyToStart) {
				Return
			}
			
			; タイマー開始
			if (this.bShowMsgs) {
				sMsg := Round(this.fIntervalMin, 1) . "分タイマーを開始します！"
				ShowAutoHideTrayTip("キッチンタイマー", sMsg, giKITCHENTIMER_TRAYTIP_DURATION_MS)
			}
			iIntervalMs := Integer(this.fIntervalMin * 60 * 1000)
			SetTimer this.objCallbackFunc, iIntervalMs
			
			; ログファイル生成
			if (this.bCreateLogFile) {
				sStartDateTime := A_Now
				sTargetDateTime := DateAdd(sStartDateTime, Integer(iIntervalMs / 1000), "Seconds")
				if (this.bIsRestart) {
					sLogStartDateTime := this.sOrigStartDateTime
					fLogIntervalMin := this.fOrigIntervalMin
				} else {
					sLogStartDateTime := sStartDateTime
					fLogIntervalMin := this.fIntervalMin
				}
				sLogFilePath :=
					EnvGet("MYDIRPATH_DESKTOP") . "\KitchenTimer_" .
					FormatTime(sLogStartDateTime, "yyyyMMdd-HHmmss") . "_" .
					Round(fLogIntervalMin, 1) . "min.log"
				sFileContents := sTargetDateTime . "`n" . fLogIntervalMin . "`n" . sLogStartDateTime
				FileAppend sFileContents, sLogFilePath
				this.sLogFilePath := sLogFilePath
			}
		} ; }}}
		Stop() { ; {{{
			; タイマークリア
			if (this.bIsRestart) {
				sMsg := Round(this.fOrigIntervalMin, 1) . "分タイマーが終了しました！"
			} else {
				sMsg := Round(this.fIntervalMin, 1) . "分タイマーが終了しました！"
			}
			ShowAutoHideTrayTip("キッチンタイマー", sMsg, giKITCHENTIMER_TRAYTIP_DURATION_MS)
			SetTimer this.objCallbackFunc, 0
			
			; ログファイル削除
			if (this.bCreateLogFile) {
				DeleteFile(this.sLogFilePath)
			}
		} ; }}}
		TimerCallback() { ; {{{
			; 画面フラッシュ
			FlashScreen()
			
			; タイマークリア
			this.Stop()
		} ; }}}
	} ; }}}

	; アラームタイマー
	SetEveryDayAlermTimer() ; {{{
	{
		iCurWeekDay := Integer(FormatTime(A_Now, "WDay")) ; 1:Sun, 2:Mon, ... 7:Sat
		bIsTargetWeekDay := ExistArrayValue(aiALARMTIMER_EVERYDAY_TRGT_WEEKDAY, iCurWeekDay)
		if (bIsTargetWeekDay) {
			SetAlermTimer("8:57", 0.0, false, false)
			SetAlermTimer("11:57", 0.0, false, false)
			SetAlermTimer("12:57", 0.0, false, false)
			SetAlermTimer("17:57", 0.0, false, false)
		}
	} ; }}}
	ClearAlermTimer() ; {{{
	{
		sLogFilePath := EnvGet("MYDIRPATH_DESKTOP") . "\AlermTimer_*.log"
		DeleteFile(sLogFilePath)
	} ; }}}
	SetAlermTimer(sTargetClock:="", fSnoozeTimeMin?, bShowMsgs:=true, bCreateLogFile:=true) ; {{{
	{
		alerm_timer := AlermTimer(bShowMsgs, bCreateLogFile)
		if (IsSet(fSnoozeTimeMin)) {
			alerm_timer.SetTimeWithClock(sTargetClock, fSnoozeTimeMin)
		} else {
			alerm_timer.SetTimeWithClock(sTargetClock)
		}
		alerm_timer.Start()
	} ; }}}
	RestartAlermTimer() ; {{{
	{
		sFilePattern := EnvGet("MYDIRPATH_DESKTOP") . "\AlermTimer_*.log"
		Loop Files, sFilePattern
		{
			sFileName := A_LoopFileName
			sFilePath := EnvGet("MYDIRPATH_DESKTOP") . "\" . sFileName
			
			alerm_timer := AlermTimer(false, true)
			alerm_timer.SetTimeWithLogFile(sFilePath)
			alerm_timer.Start()
		}
	} ; }}}
	class AlermTimer { ; {{{
		__New(bShowMsgs:=true, bCreateLogFile:=true) { ; {{{
			this.objCallbackFunc := ObjBindMethod(this, "TimerCallback")
			this.bShowMsgs := bShowMsgs
			this.bCreateLogFile := bCreateLogFile
			this.sTargetDateTime := ""
			this.bIsRestart := false
			this.bReadyToStart := false
			this.fSnoozeTimeMin := 0.0
		} ; }}}
		SetTimeWithLogFile(sLogFilePath) { ; {{{
			this.bIsRestart := true
			if (!FileExist(sLogFilePath)) {
				MsgBox "ファイルが存在しません。`n  " . sLogFilePath . "`n`n処理を中断します。"
				Return
			}
			sFileLines := FileRead(sLogFilePath)
			asFileLines := StrSplit(sFileLines, "`n")
			sTargetDateTime := asFileLines[1]
			sOrigStartDateTime := asFileLines[2]
			fSnoozeTimeMin := asFileLines[3]
			;MsgBox sTargetDateTime . "`n" . sOrigStartDateTime . "`n" . fSnoozeTimeMin
			
			this.sOrigStartDateTime := sOrigStartDateTime
			
			DeleteFile(sLogFilePath)
			
			sStartDateTime := A_Now
			iSecond := DateDiff(sTargetDateTime, sStartDateTime, "Seconds")
			if (iSecond > 0) {
				this.sTargetDateTime := sTargetDateTime
				this.fSnoozeTimeMin := fSnoozeTimeMin
				this.bReadyToStart := true
			}
		} ; }}}
		SetTimeWithClock(sTargetClock:="", fSnoozeTimeMin?) { ; {{{
			this.bIsRestart := false
			sCurDateTime := A_Now
			; アラーム時刻設定
			if (sTargetClock == "") {
				; 時刻初期値生成
				sInitTime := ""
				if (gbALARMTIMER_INITTIME_CUR) {
					sInitTime := FormatTime(sCurDateTime, "HH:mm")
				} else {
					iCurHour := Integer(FormatTime(sCurDateTime, "HH"))
					iCurMinutes := Integer(FormatTime(sCurDateTime, "mm"))
					sInitHour := ""
					sInitMinute := ""
					if (
					  (giALARMTIMER_INITTIME_MIN_STEP <= 0) ||
					  (giALARMTIMER_INITTIME_MIN_STEP > 60) ||
					  (Mod(60, giALARMTIMER_INITTIME_MIN_STEP) != 0)
					) {
						MsgBox "[fatal error] 時刻初期値設定に誤りがあるため、処理を中断します。"
						Return
					}
					iMinTrgt := 0
					while iCurMinutes >= iMinTrgt
					{
						iMinTrgt := iMinTrgt + giALARMTIMER_INITTIME_MIN_STEP
					}
					if (iMinTrgt = 60) {
						if (iCurHour < 23) {
							sInitHour := iCurHour + 1
						} else {
							sInitHour := "00"
						}
						sInitMinute := "00"
					} else {
						sInitHour := iCurHour
						sInitMinute := Format("{1:02d}" , String(iMinTrgt))
					}
					sInitTime := sInitHour . ":" . sInitMinute
				}
				
				; 時刻入力
				InputBoxObj := InputBox("アラームを設定します。`n時刻（e.g. 12:30）を設定してください。", "アラームタイマー", , sInitTime)
				if (InputBoxObj.Result = "Cancel") {
					MsgBox "キャンセルされたため、処理を中断します。"
					Return
				}
				sTargetClock := InputBoxObj.Value
			}
			; アラーム時刻フォーマットチェック
			Try {
				iDelimiterPos := Integer(InStr(sTargetClock, ":"))
				;MsgBox String(iDelimiterPos)
				if (iDelimiterPos = 0) { ; delimiter not found
					throw
				}
				iTargetHour := Integer(SubStr(sTargetClock, 1, iDelimiterPos-1))
				iTargetMinutes := Integer(SubStr(sTargetClock, iDelimiterPos+1, 2))
				;MsgBox iTargetHour . " " . iTargetMinutes
				if (iTargetHour < 0 || iTargetHour > 23) {
					throw
				}
				if (iTargetMinutes < 0 || iTargetMinutes > 59) {
					throw
				}
			} Catch Error as err {
				MsgBox "不正な時刻が指定されました。`n" . sTargetClock . "`n`n処理を中断します。"
				Return
			}
			sTargetDateTime := A_YYYY . A_MM . A_DD . Format("{1:02d}" , String(iTargetHour)) . Format("{1:02d}" , String(iTargetMinutes)) . "00"
			if (DateDiff(sTargetDateTime, sCurDateTime, "Seconds") < 0) {
				sTargetDateTime := DateAdd(sTargetDateTime, 1, "days")
			}
			this.sTargetDateTime := sTargetDateTime
			
			; スヌーズ時間設定
			if (!IsSet(fSnoozeTimeMin)) {
				;スヌーズ時間入力
				InputBoxObj := InputBox("スヌーズ時間を設定します。`n0.0より大きい値（e.g. 0.5）を設定してください。", "アラームタイマー", , 0.0)
				if (InputBoxObj.Result = "Cancel") {
					MsgBox "キャンセルされたため、処理を中断します。"
					Return
				}
				sSnoozeTimeMin := InputBoxObj.Value
				Try {
					fSnoozeTimeMin := Float(sSnoozeTimeMin)
				} Catch Error as err {
					MsgBox "不正なスヌーズ時間が指定されました。`n" . sSnoozeTimeMin . "`n`n処理を中断します。"
					Return
				}
			}
			; スヌーズ時間フォーマットチェック
			Try {
				if (fSnoozeTimeMin < 0.0) {
					throw
				}
			} Catch Error as err {
				MsgBox "不正なスヌーズ時間が指定されました。`n" . sSnoozeTimeMin . "`n`n処理を中断します。"
				Return
			}
			this.fSnoozeTimeMin := fSnoozeTimeMin
			
			this.bReadyToStart := true
		} ; }}}
		SetTimeWithTargetDateTime(sTargetDateTime) { ; {{{
			sTargetYear := FormatTime(sTargetDateTime, "yyyy")
			if (sTargetYear = "") {
				MsgBox "不正な時刻が指定されました。`n" . sTargetDateTime . "`n`n処理を中断します。"
				Return
			}
			this.sTargetDateTime := sTargetDateTime
			this.bReadyToStart := true
		} ; }}}
		SetSnoozeTime(fSnoozeTimeMin) { ; {{{
			if (!isFloat(fSnoozeTimeMin)) {
				Return
			}
			this.fSnoozeTimeMin := fSnoozeTimeMin
		} ; }}}
		Start() { ; {{{
			; sTargetDateTime: target date time. this time must be YYYYMMDDHH24MISS format.
			; 事前チェック
			if (!this.bReadyToStart) {
				Return
			}
			sTargetDateTime := this.sTargetDateTime
			fSnoozeTimeMin := this.fSnoozeTimeMin
			
			sStartDateTime := A_Now
			iIntervalMs := DateDiff(sTargetDateTime, sStartDateTime, "Seconds") * 1000
			iIntervalMin := DateDiff(sTargetDateTime, sStartDateTime, "Minutes")
			;MsgBox iIntervalMs . "`n" . sStartDateTime
			
			; タイマー開始
			sIntervalTime := ""
			if (iIntervalMin >= 60) {
				sIntervalTime := Format("{1:d}" , iIntervalMin / 60) . "時間" . String(Mod(iIntervalMin, 60)) . "分"
			} else {
				sIntervalTime := iIntervalMin . "分"
			}
			
			if (this.bShowMsgs) {
				sMsg := FormatTime(sTargetDateTime, "HH:mm") . "（" . sIntervalTime . "後）にアラームを設定しました！"
				ShowAutoHideTrayTip("アラームタイマー", sMsg, giALARMTIMER_TRAYTIP_DURATION_MS)
			}
			SetTimer this.objCallbackFunc, iIntervalMs
			
			; ログファイル生成
			if (this.bCreateLogFile) {
				if (this.bIsRestart) {
					sLogStartDateTime := this.sOrigStartDateTime
				} else {
					sLogStartDateTime := sStartDateTime
				}
				sLogFilePath :=
					EnvGet("MYDIRPATH_DESKTOP") . "\AlermTimer_" .
					FormatTime(sLogStartDateTime, "yyyyMMdd-HHmmss") .  "_to_" .
					FormatTime(sTargetDateTime, "yyyyMMdd-HHmmss") . ".log"
				sContents := sTargetDateTime . "`n" . sLogStartDateTime . "`n" . fSnoozeTimeMin
				FileAppend sContents, sLogFilePath
				this.sLogFilePath := sLogFilePath
			}
		} ; }}}
		TimerCallback() { ; {{{
			; 画面フラッシュ
			FlashScreen()
			
			; タイマー満了通知
			sMsg := FormatTime(this.sTargetDateTime, "HH:mm") . "になりました！"
			ShowAutoHideTrayTip("アラームタイマー", sMsg, giALARMTIMER_TRAYTIP_DURATION_MS)
			
			if (this.fSnoozeTimeMin > 0.0) {
				sAnswer := MsgBox("スヌーズを停止しますか？", "アラームタイマー", "Y/N T" . String(giALARMTIMER_SNOOZE_MSG_DURATION_SEC) . " Default2")
				if (sAnswer == "Yes") {
					; タイマークリア
					sMsg := this.fSnoozeTimeMin . "分間のスヌーズを停止しました！"
					ShowAutoHideTrayTip("アラームタイマー", sMsg, giALARMTIMER_TRAYTIP_DURATION_MS)
					SetTimer this.objCallbackFunc, 0
					
					; ログファイル削除
					if (this.bCreateLogFile) {
						DeleteFile(this.sLogFilePath)
					}
				} else { ; sAnswer == "No" or "Timeout"
					sMsg := this.fSnoozeTimeMin . "分間のスヌーズを設定しました！"
					ShowAutoHideTrayTip("アラームタイマー", sMsg, giALARMTIMER_TRAYTIP_DURATION_MS)
					
					; タイマー再設定
					iIntervalMs := Integer(this.fSnoozeTimeMin * 60 * 1000)
					SetTimer this.objCallbackFunc, iIntervalMs
				}
			} else {
				; タイマークリア
				SetTimer this.objCallbackFunc, 0
				
				; ログファイル削除
				if (this.bCreateLogFile) {
					DeleteFile(this.sLogFilePath)
				}
			}
		} ; }}}
	} ; }}}

;* ***************************************************************
;* Functions (common)
;* ***************************************************************
	; 値のクリッピング
	CropValue(iValue, iMin, iMax) ; {{{
	{
		if ( iValue < iMin ) {
			return iMin
		} else {
			if ( iValue > iMax ) {
				return iMax
			} else {
				return iValue
			}
		}
	} ; }}}
	; 配列内の値存在確認
	ExistArrayValue(axArray, xValue) ; {{{
	{
		bIsExist := false
		Loop (axArray.Length) {
			if (xValue == axArray[A_Index]) {
				bIsExist := true
			}
		}
		return bIsExist
	} ; }}}
	; ファイル削除
	DeleteFile(sFilePattern) ; {{{
	{
		if (FileExist(sFilePattern)) {
			FileDelete sFilePattern
		}
	} ; }}}

