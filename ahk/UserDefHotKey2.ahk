;	[Help] https://www.autohotkey.com/docs/v2/index.htm

;	#NoTrayIcon						; スクリプトのタスクトレイアイコンを非表示にする。
	#Warn All						; Enable warnings to assist with detecting common errors.
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え
	#WinActivateForce				; ウィンドウのアクティブ化時に、穏やかな方法を試みるのを省略して常に強制的な方法でアクティブ化を行う。（タスクバーアイコンが点滅する現象が起こらなくなる）
	SendMode "Input"				; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。

;* ***************************************************************
;* Settings
;* ***************************************************************
global gsDOC_DIR_PATH := "C:\Users\" . A_Username . "\Dropbox\100_Documents"
global giWIN_TILE_MODE_CLEAR_INTERVAL := 10000 ; [ms]
global giWIN_TILE_MODE_MAX := 3 ; 0～giWIN_TILE_MODE_MAX
global giWIN_Y_OFFSET := 2/7
global giWIN_TILE_MODE_OFFSET := 0
global giSCREEN_BRIGHTNESS_STEP := 20 ; 0～100 [%]
global giSCREEN_BRIGHTNESS_MIN := giSCREEN_BRIGHTNESS_STEP ; 0～100 [%]
global giSCREEN_BRIGHTNESS_MAX := 100 ; 0～100 [%]
global giSCREEN_BRIGHTNESS_INIT := giSCREEN_BRIGHTNESS_MAX
global giSTART_PRG_TOOL_TIP_SHOW_TIME := 2000 ; [ms]

;* ***************************************************************
;* Preprocess
;* ***************************************************************
TraySetIcon "UserDefHotKey2.ico"
ShowAutoHideTrayTip(A_ScriptName, A_ScriptName . " is loaded.", 2000)
InitScreenBrightness()
InitWinTileMode()
GetCurYearMonths()

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
		VK1D & Right::	SendKeyWithModKeyCurPressing( "End" )
		VK1D & Left::	SendKeyWithModKeyCurPressing( "Home" )
		VK1D & Up::		SendKeyWithModKeyCurPressing( "PgUp" )
		VK1D & Down::	SendKeyWithModKeyCurPressing( "PgDn" )
		Insert::Return																												;Insertキー
		PrintScreen::return																											;PrintScreenキー

;***** ホットキー（Global） *****
	;スクリプトリロード
		^+!F5::		ReloadMe()
	;ファイルオープン
		^+!a::		StartProgramAndActivate( EnvGet("MYEXEPATH_GVIM"), A_ScriptFullPath )											;UserDefHotKey.ahk
		^+!Space::	StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#temp.txt" )													;#temp.txt
		^+!Down::	StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#temp.txt" )													;#temp.txt
		^+!Enter::																													;#todo.itmz
		^+!Up::																														;#todo.itmz
		{
		;	lPID := ProcessWait("Dropbox.exe", 30) ; Dropboxが起動(≒同期が完了)するまで待つ(タイムアウト時間30s)
			StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#todo.itmz" )
		}
		^+!Right::	StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#temp.xlsm" )													;#temp.xlsm
		^+!Left::	StartProgramAndActivateFile( gsDOC_DIR_PATH . "\#temp.vsdm" )													;#temp.vsdm
		^+!F1::		StartProgramAndActivateFile( "C:\other\グローバルホットキー配置.vsdx" )											;ホットキー配置表示
		^+!\::		StartProgramAndActivateFile( gsDOC_DIR_PATH . "\210_【衣食住】家計\100_予算管理.xlsm" )							;予算管理.xlsm
		^+!^::		StartProgramAndActivateFile( gsDOC_DIR_PATH . "\..\000_Public\家計\予算管理＠家族用.xlsx" )						;予算管理＠家族用.xlsx
		^+!/::		StartProgramAndActivateFile( gsDOC_DIR_PATH . "\320_【自己啓発】勉強\words.itmz" )								;用語集
		^+!c::		StartProgramAndActivateFile( "C:\other\言語チートシート.xlsx" )													;言語チートシート
		^+!s::		StartProgramAndActivateFile( "C:\other\ショートカットキー一覧.xlsx" )											;ショートカットキー
		^+!o::		StartProgramAndActivateFile( "C:\other\template\#object.xlsm" )													;#object.xlsm
	;プログラム起動
		^+!y::		StartProgramAndActivateFile( EnvGet("MYDIRPATH_CODES") . "\_sync_github-codes-remote.bat" )						;codes同期
		^+!k::		StartProgramAndActivateFile( EnvGet("MYDIRPATH_CODES") . "\vbs\tools\win\other\KitchenTimer.vbs" )				;KitchenTimer.vbs
		^+!t::		StartProgramAndActivateFile( EnvGet("MYDIRPATH_CODES") . "\vbs\tools\win\other\PeriodicKeyTransmission.bat" )	;定期キー送信
		^+!w::		StartProgramAndActivateFile( EnvGet("MYDIRPATH_CODES") . "\vbs\tools\win\file_ope\CopyRefFileFromWeb.vbs" )		;Webから参照ファイル取得
		^+!;::		StartProgramAndActivateExe( EnvGet("MYEXEPATH_CALC"), True )													;cCalc.exe
		^+!x::																														;rapture.exe
		{
			SetBrightnessTemporary(giSCREEN_BRIGHTNESS_MAX, 5000)
			StartProgramAndActivateExe( EnvGet("MYEXEPATH_RAPTURE"), False, False )
		}
	;フォルダ表示
		^+!z::																														;ファイラ―
		{
		;	;xf.exe
		;	StartProgramAndActivateExe( EnvGet("MYEXEPATH_XF"), 1 )
			;エクスプローラー
			StartProgramAndActivateFile( gsDOC_DIR_PATH )
			Sleep 100
			Send "+{tab}"
		}
		^+!F12::																													;Programsフォルダ表示
		{
			StartProgramAndActivateFile( "C:\Users\" . A_Username . "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs" )
			Sleep 100
			Send "+{tab}"
		}
	;サイトオープン
		^+!1::	Run "https://draemonash2.github.io/"																				;Github.io
		^+!2::	Run "https://draemonash2.github.io/linux_sft/linux.html"															;Github.io linux
		^+!3::	Run "https://draemonash2.github.io/gitcommand_lng/gitcommand.html"													;Github.io git command
		^+!h::																														;翻訳サイト
		{
		;	Run "https://translate.google.com/?sl=en&tl=ja&op=translate&hl=ja"
			Run "https://www.deepl.com//translator"
		}
	;Wifi接続
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
	;Window最前面化
		^!Pause::	; HP製PC以外：Ctrl+Alt+Pause、HP製PC：Ctrl+Shift+Alt+Fn（HP製PCでは「Pause」＝「Fn＋Shift」）
		{
			static bEnableAlwaysOnTop := 0
			WinSetAlwaysOnTop -1, "A"
			sActiveWinTitle := WinGetTitle("A")
			if (bEnableAlwaysOnTop = 0)
			{
				MsgBox "Window最前面を【有効】にします`n`n" . sActiveWinTitle, "Window最前面化", 0x43000
				bEnableAlwaysOnTop := 1
			}
			else
			{
				MsgBox "Window最前面を【解除】します`n`n" . sActiveWinTitle, "Window最前面化", 0x43000
				bEnableAlwaysOnTop := 0
			}
		}
	;Windowタイル切り替え
		!#LEFT::
		{
			;MsgBox "!#LEFT"
			SetTimerWinTileMode()
			IncrementWinTileMode()
			ApplyWinTileMode()
		}
		!#RIGHT::
		{
			;MsgBox "!#RIGHT"
			SetTimerWinTileMode()
			DecrementWinTileMode()
			ApplyWinTileMode()
		}
	;画面明るさ設定
		#Home::	SetBrightness(giSCREEN_BRIGHTNESS_MAX)
		#End::	SetBrightness(giSCREEN_BRIGHTNESS_MIN)
		#PgDn::	DarkenScreen()
		#PgUp::	BrightenScreen()
	;テスト用
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

;***** ホットキー(Software local) *****
	#HotIf !WinActive("ahk_exe WindowsTerminal.exe")
		RAlt::Send "{AppsKey}"	;右Altキーをコンテキストメニュー表示に変更
	#HotIf
	
	#HotIf WinActive("ahk_exe explorer.exe")
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
		^+g::	Run EnvGet("MYEXEPATH_TRESGREP") . " " . GetCurDirPathAtExplorer()													; Grep検索＠TresGrep
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
	; 起動＆アクティベート処理 (実行プログラム＆ファイルパス指定)
	StartProgramAndActivate( sExePath, sFilePath, bLaunchSingleProcess:=False, bShowToolTip:=True )
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
			ShowAutoHideToolTip(sFileName . " is starting...", giSTART_PRG_TOOL_TIP_SHOW_TIME)
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
		} Catch Error as err {
			MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
				, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
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
	StartProgramAndActivateFile( sFilePath, bShowToolTip:=True )
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
			ShowAutoHideToolTip(sFileName . " is starting...", giSTART_PRG_TOOL_TIP_SHOW_TIME)
		}
		
		;*** start program ***
		Try {
			Run sFilePath, , , &sOutputVarPID
		} Catch Error as err {
			MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
				, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
		}
	;	WinActivate "ahk_pid " . sOutputVarPID
		return
	}
	
	; 起動＆アクティベート処理 (実行プログラム指定のみ)
	StartProgramAndActivateExe( sExePath, bLaunchSingleProcess:=False, bShowToolTip:=True )
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
			ShowAutoHideToolTip(sExeName . " is starting...", giSTART_PRG_TOOL_TIP_SHOW_TIME)
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
		} Catch Error as err {
			MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
				, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
			return
		}
	;	WinActivate "ahk_pid " . sOutputVarPID
		return
	}

	; 今押している修飾キーと共にキー送信する
	SendKeyWithModKeyCurPressing( sSendKey )
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
	}

	;Windowタイル切り替え
	InitWinTileMode()
	{
		global giWinTileMode := GetWinTileModeMin()
		SetTimerWinTileMode()
	}
	IncrementWinTileMode()
	{
		global giWinTileMode
		iWinTileModeMin := GetWinTileModeMin()
		if ( giWinTileMode >= giWIN_TILE_MODE_MAX ) {
			giWinTileMode := iWinTileModeMin
		} else {
			giWinTileMode := giWinTileMode + 1
		}
	;	MsgBox "[DBG] IncrementWinTileMode()" . "`ngiWinTileMode = " . giWinTileMode . "`niWIN_TILE_MODE_MAX = " . giWIN_TILE_MODE_MAX . "`niWinTileModeMin = " . iWinTileModeMin
	}
	DecrementWinTileMode()
	{
		global giWinTileMode
		iWinTileModeMin := GetWinTileModeMin()
		if ( giWinTileMode <= iWinTileModeMin ) {
			giWinTileMode := giWIN_TILE_MODE_MAX
		} else {
			giWinTileMode := giWinTileMode - 1
		}
	;	MsgBox "[DBG] DecrementWinTileMode()" . "`ngiWinTileMode = " . giWinTileMode . "`n giWIN_TILE_MODE_MAX = " . giWIN_TILE_MODE_MAX . "`niWinTileModeMin = " . iWinTileModeMin
	}
	; ウィンドウサイズ切り替え
	ApplyWinTileMode()
	{
		global giWinTileMode
		GetMonitorPosInfo(1, &mainx, &mainy, &mainwidth, &mainheight )
		iMonitorNum := GetMonitorNum()
		if (iMonitorNum > 1) {
			GetMonitorPosInfo(2, &subx, &suby, &subwidth, &subheight )
			subywhole := Integer(suby + ( subheight * giWIN_Y_OFFSET ))
			subheightwhole := Integer(subheight * ( 1 - giWIN_Y_OFFSET ))
		}
		switch giWinTileMode
		{
			case 0:	;サブ全体
				winx		:= subx
				winy		:= subywhole
				winwidth	:= subwidth
				winheight	:= subheightwhole
			case 1:	;サブ上
				winx		:= subx
				winy		:= subywhole
				winwidth	:= subwidth
				winheight	:= Integer(subheightwhole / 2)
			case 2:	;サブ下
				winx		:= subx
				winy		:= Integer(subywhole + subheightwhole / 2)
				winwidth	:= subwidth
				winheight	:= Integer(subheightwhole / 2)
			case 3:	;メイン全体
				winx		:= mainx
				winy		:= mainy
				winwidth	:= mainwidth
				winheight	:= mainheight
			case 4:	;メイン左
				winx		:= Integer(mainx - giWIN_TILE_MODE_OFFSET)
				winy		:= mainy
				winwidth	:= Integer(mainwidth / 2 + giWIN_TILE_MODE_OFFSET)
				winheight	:= Integer(mainheight + giWIN_TILE_MODE_OFFSET)
			case 5:	;メイン右
				winx		:= mainx + Integer(mainwidth / 2 - giWIN_TILE_MODE_OFFSET)
				winy		:= mainy
				winwidth	:= Integer(mainwidth / 2 + giWIN_TILE_MODE_OFFSET)
				winheight	:= Integer(mainheight + giWIN_TILE_MODE_OFFSET)
			default:
				MsgBox "[error] invalid giWinTileMode.`n" . giWinTileMode
				return
		}
	;	MsgBox "[DBG] ApplyWinTileMode() " .
	;		"`ngiWinTileMode = " . giWinTileMode .
	;		"`nmainx = " . mainx . "`nmainy = " . mainy . "`nmainwidth = " . mainwidth . "`nmainheight = " . mainheight . "`nsubx = " . subx . "`nsuby = " . suby . "`nsubwidth = " . subwidth . "`nsubheight = " . subheight .
	;		"`nsubywhole = " . subywhole . "`nsubheightwhole = " . subheightwhole .
	;		"`nwinx = " . winx . "`nwiny = " . winy . "`nwinwidth = " . winwidth . "`nwinheight = " . winheight
		WinMove winx, winy, winwidth, winheight, "A"
		return
	}
	GetWinTileModeMin()
	{
		iMonitorNum := GetMonitorNum()
		if (iMonitorNum = 2) {
			iWinTileModeMin := 0
		} else {
			iWinTileModeMin := 3
		}
	;	MsgBox "[DBG] IncrementWinTileMode()" . "`niMonitorNum = " . iMonitorNum . "`niWinTileModeMin = " . iWinTileModeMin
		return iWinTileModeMin
	}
	GetMonitorNum()
	{
		return SysGet(80) ; SM_CMONITORS: Number of display monitors on the desktop (not including "non-display pseudo-monitors").
	}
	GetMonitorPosInfo( MonitorNum, &X, &Y, &Width, &Height )
	{
		try
		{
			ActualN := MonitorGetWorkArea(MonitorNum, &Left, &Top, &Right, &Bottom)
		;	MsgBox "Left: " Left " -- Top: " Top " -- Right: " Right " -- Bottom: " Bottom
		} Catch Error as err {
			MsgBox Format("{1}: {2}.`n`nFile:`t{3}`nLine:`t{4}`nWhat:`t{5}`nStack:`n{6}"
				, type(err), err.Message, err.File, err.Line, err.What, err.Stack)
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
	;	MsgBox "[DBG] GetMonitorPosInfo()" . "`nMonitorNum = " . MonitorNum . "`nX = " . X . "`nY = " . Y . "`nWidth = " . Width . "`nHeight = " . Height
	}
	SetTimerWinTileMode()
	{
		SetTimer ClearWinTileMode, giWIN_TILE_MODE_CLEAR_INTERVAL
	}
	ClearWinTileMode()
	{
		global giWinTileMode
		giWinTileMode := giWIN_TILE_MODE_MAX
		;ShowAutoHideTrayTip("タイルモードクリアタイマー", "タイルモードをクリアしました", 5000)
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
	;	MsgBox "sTrgtPaths = " . sTrgtPaths
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
	;	MsgBox "sTrgtPaths = " . sTrgtPaths
		return sTrgtPaths
	}
	; 選択ファイル名取得＠explorer
	GetSelFileNameAtExplorer()
	{
		sFilePaths := GetSelFilePathAtExplorer(0)
		sDirPaths := GetCurDirPathAtExplorer()
		sTrgtPaths := StrReplace(sFilePaths, sDirPaths . "\", )
		sTrgtPaths := StrReplace(sTrgtPaths, "`"", )
	;	MsgBox "sTrgtPaths = " . sTrgtPaths . "`nsFilePaths = " . sFilePaths . "`nsDirPaths = " . sDirPaths
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

	; ツールチップ表示
	ShowAutoHideToolTip(sMsg, iShowPeriodMs)
	{
		ToolTip(sMsg)
		SetTimer(HideToolTip, -1 * iShowPeriodMs)
		Return
	}
	HideToolTip()
	{
		ToolTip()
		Return
	}

	; トレイチップ表示
	ShowAutoHideTrayTip(sTitle, sMsg, iShowPeriodMs)
	{
		TrayTip sMsg, sTitle, 1
		SetTimer(HideTrayTip, -1 * iShowPeriodMs)
		Return
	}
	HideTrayTip()
	{
		TrayTip()
		Return
	}

	; 画面明るさ設定
	InitScreenBrightness()
	{
		global giBrightness := giSCREEN_BRIGHTNESS_INIT
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
	}
	SetBrightness(iBrightness)
	{
		global giBrightness
		giBrightness := iBrightness
		ApplyBrightness()
		ShowAutoHideToolTip("明るさ：" . giBrightness . "%", 500)
	}
	BrightenScreen()
	{
		global giBrightness
		giBrightness += giSCREEN_BRIGHTNESS_STEP
		if (giBrightness > giSCREEN_BRIGHTNESS_MAX)
		{
			giBrightness := giSCREEN_BRIGHTNESS_MAX
		}
		ApplyBrightness()
		ShowAutoHideToolTip("明るさ：" . giBrightness . "%", 500)
	}
	DarkenScreen()
	{
		global giBrightness
		giBrightness -= giSCREEN_BRIGHTNESS_STEP
		if (giBrightness < giSCREEN_BRIGHTNESS_MIN)
		{
			giBrightness := giSCREEN_BRIGHTNESS_MIN
		}
		ApplyBrightness()
		ShowAutoHideToolTip("明るさ：" . giBrightness . "%", 500)
	}
	SetBrightnessTemporary(iBrightness, iWaitTimeMs)
	{
		global giBrightness
		global giBrightnessOld := giBrightness
		giBrightness := iBrightness
		ApplyBrightness()
		SetTimer(SetOldBrightness, -1 * iWaitTimeMs)
	}
	SetOldBrightness()
	{
		global giBrightness
		global giBrightnessOld
		giBrightness := giBrightnessOld
		ApplyBrightness()
	}
	ApplyBrightness()
	{
		global giBrightness
		iMonitorCount := MonitorGetCount()
		Loop iMonitorCount
		{
			iDimId := gasDimId[A_Index]
			iTransparency := 100 - giBrightness
			WinSetTransparent(Integer(iTransparency * 255 / 100), "ahk_id " . iDimId)
		}
		Return
	}

	;クリップボード設定
	SetClipboard(sStr)
	{
		A_Clipboard := ""
		A_Clipboard := sStr
		ClipWait
	}

	; GUI
	CreateSlctCmndWindowZip()
	{
		global myGui
		global ogcListBoxAnswer
		myGui := Gui()
		myGui.Add("Text", , "圧縮/パスワード圧縮/解凍を実行します。`n処理を選択してください。")
		ogcListBoxAnswer := myGui.Add("ListBox", "vAnswer Choose1 R3", ["圧縮", "パスワード圧縮", "解凍"])
		ogcButtonZipEnter := myGui.Add("Button", "Hidden w0 h0 Default", "ZipEnter")
		ogcButtonZipEnter.OnEvent("Click", EventClickAtZip.Bind("Normal"))
		myGui.OnEvent("Close", EventEscape)
		myGui.OnEvent("Escape", EventEscape)
		myGui.Show("Center")
	}
	EventClickAtZip(A_GuiEvent, GuiCtrlObj, Info, *)
	{
		global myGui
		global ogcListBoxAnswer
		vAnswer := ogcListBoxAnswer.Text
		;MsgBox vAnswer
		myGui.Destroy()
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
	}
	CreateSlctCmndWindowLink()
	{
		global myGui
		global ogcListBoxAnswer
		myGui := Gui()
		myGui.Add("Text", , "ショートカット/シンボリックリンクを作成します。`n処理を選択してください。")
		ogcListBoxAnswer := myGui.Add("ListBox", "vAnswer Choose1 R2", ["ショートカット作成", "シンボリックリンク作成"])
		ogcButtonSelLinkEnter := myGui.Add("Button", "Hidden w0 h0 Default", "SelLinkEnter")
		ogcButtonSelLinkEnter.OnEvent("Click", EventClickAtLink.Bind("Normal"))
		myGui.OnEvent("Close", EventEscape)
		myGui.OnEvent("Escape", EventEscape)
		myGui.Show("Center")
	}
	EventClickAtLink(A_GuiEvent, GuiCtrlObj, Info, *)
	{
		global myGui
		global ogcListBoxAnswer
		vAnswer := ogcListBoxAnswer.Text
		myGui.Destroy()
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
	}
	CreateSlctCmndWindowPathList()
	{
		global myGui
		global ogcListBoxAnswer
		myGui := Gui()
		myGui.Add("Text", , "パス一覧を作成します。`n処理を選択してください。")
		ogcListBoxAnswer := myGui.Add("ListBox", "vAnswer Choose1 R4", ["ファイル＆フォルダ一覧作成", "ファイル一覧作成", "フォルダ一覧作成", "フォルダツリー作成"])
		ogcButtonPathListEnter := myGui.Add("Button", "Hidden w0 h0 Default", "PathListEnter")
		ogcButtonPathListEnter.OnEvent("Click", EventClickPathList.Bind("Normal"))
		myGui.OnEvent("Close", EventEscape)
		myGui.OnEvent("Escape", EventEscape)
		myGui.Show("Center")
	}
	EventClickPathList(A_GuiEvent, GuiCtrlObj, Info, *)
	{
		global myGui
		global ogcListBoxAnswer
		vAnswer := ogcListBoxAnswer.Text
		myGui.Destroy()
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
	}
	EventEscape(*)
	{
		global myGui
	;	MsgBox "エスケープされました"
		myGui.Destroy()
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

	; 本スクリプトをリロードする
	ReloadMe()
	{
		Reload
		Sleep 1000 ; リロードに成功した場合、リロードはスリープ中にこのインスタンスを閉じるので、以下の行に到達することはない
		MsgBox "スクリプト" . A_ScriptName . "の再読み込みに失敗しました"
	}

	; 今月/先月の月日を取得する
	GetCurYearMonths()
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
	}
	; 月日を取得する
	GetYearMonth(sInYear, sInMonth, &sOutYear, &sOutMonth1Deg, &sOutMonth2Deg, iOffset:=0 )
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
	}
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

