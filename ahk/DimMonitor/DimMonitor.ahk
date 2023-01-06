; DimMonitor.ahk 2.1.0
; https://sites.google.com/site/bucuerider/

#NoEnv
#SingleInstance force

Dim = 0
DMPause = 0
SplitPath, A_LineFile, , LineDir
DimIni := LineDir "\DimMonitor.ini"
If !FileExist(DimIni)
{
	FileAppend,
(
[General]
;時刻による自動調整（0：しない　1：する）
AutoDim=0
;自動調整を日の出・日の入り時刻に合わせて自動調整（0：しない　1：する）
SunTime=1
;自動調整で明るくする時・分
DayH=6
DayM=0
;自動調整で暗くする時・分
NightH=18
NightM=0
;手動調整後に自動調整を停止する間隔（時間単位）
Span=2
;自動調整時の明るさ（20-100％）
AutoSet=75
;フィルタの色（HTMLカラーコード参照）
DimColor=000000
), % DimIni
}
IniRead, AutoDim, % DimIni, General, AutoDim, 1
IniRead, SunTime, % DimIni, General, SunTime, 0
If (AutoDim == 1 && SunTime != 1)
{
	IniRead, DayH, % DimIni, General, DayH, 6
	IniRead, DayM, % DimIni, General, DayM, 0
	DayH := (DayH < 0 || DayH >= 24) ? 6 : DayH
	DayM := (DayM < 0 || DayM >= 60) ? 0 : DayM
	Day := DayH * 100 + DayM
	IniRead, NightH, % DimIni, General, NightH, 18
	IniRead, NightM, % DimIni, General, NightM, 0
	NightH := (NightH < 0 || NightH >= 24) ? 18 : NightH
	NightM := (NightM < 0 || NightM >= 60) ? 0 : NightM
	Night := NightH * 100 + NightM
}
Else If (AutoDim == 1 && SunTime == 1)
{
	Day := SunTime(0)
	Night := SunTime(1)
}
IniRead, Span, % DimIni, General, Span, 2
Span := (Span > 0) ? Span * 60 * 60 * 1000 : 0
IniRead, AutoSet, % DimIni, General, AutoSet, 75
AutoSet := (AutoSet < 20 || AutoSet >100) ? 100 : AutoSet
IniRead, DimColor, % DimIni, General, DimColor, 000000
DimColor := (RegExMatch(DimColor, "[^0-9a-fA-F]") != 0) ? 000000 : DimColor

;*** トレイメニュー作成 ***
If (A_IsCompiled || A_ScriptName = "DimMonitor.ahk")
{
	Menu, Tray, NoStandard
	Menu, Tray, Add, 明るさを指定, DMSet
	Menu, Tray, Add, 自動調整, DMAuto
	If AutoDim
		Menu, Tray, Check, 自動調整
	Menu, Tray, Add, 停止, DMPause
	Menu, Tray, Add, 終了, DMExit
	Menu, Tray, Default, 明るさを指定
	Menu, Tray, Icon, DimMonitor.ico
}
Else
{
	Menu, DMTray, Add, 明るさを指定, DMSet
	Menu, DMTray, Add, 自動調整, DMAuto
	If AutoDim
		Menu, DMTray, Check, 自動調整
	Menu, DMTray, Add, 停止, DMPause
	Menu, Tray, Add, 
	Menu, Tray, Add, DimMonitor, :DMTray
}

;*** タイマー ***
DMTimer:
SetTimer, DimOnTop, 50			; 50msごとにアクティブウィンドウの変化を監視

;*** フィルター生成 ***
DimFilter:
	SysGet, MonitorCount, MonitorCount
	Loop, %MonitorCount%
	{
		SysGet, Monitor, Monitor, %A_Index%
		Width := MonitorRight - MonitorLeft
		Height := MonitorBottom - MonitorTop
		Gui, DimGui%A_Index%:+LastFound +ToolWindow -Disabled -SysMenu -Caption +E0x20 +AlwaysOnTop
		Gui, DimGui%A_Index%:Color, %DimColor%
		Gui, DimGui%A_Index%:Show, X%MonitorLeft% Y%MonitorTop% W%Width% H%Height%, DimMonitor%A_Index%
		WinGet, DimId%A_Index%, Id, DimMonitor%A_Index% ahk_class AutoHotkeyGUI
		DimId := DimId%A_Index%
		WinSet, Transparent, % Dim * 255 / 100, ahk_id %DimId%
	}
	If (A_ScriptName = "DimMonitor.ahk" || A_ThisLabel = "DMTimer")
		Return					; 単独実行なら自動実行終了
	Else
		GoTo, DimEnd			; 組み込みならReturnしない

DimOnTop:						; 常にフィルターを最前面に配置
	IfWinNotActive, ahk_id %AWinId%
	{
		WinSet, AlwaysOnTop, On, ahk_id %DimId%
		WinGet, AWinId, Id, A
	}
	Menu, Tray, Tip, % "現在の明るさ：" 100 - Dim "%"
	GoSub, DimTime
Return

; *** トレイメニュー動作 ***
DMSet:
	InputBox, InputDim, 明るさを指定, 明るさを20から100の整数で指定してください。, , 350, 150, , , , 30, % 100 - Dim
	If (ErrorLevel != 0 || RegExMatch(InputDim, "[^0-9]") != 0)
		Return
	Else
	{
		Dim := (InputDim < 20) ? 20 : InputDim
		Dim := (Dim > 100) ? 100 : Dim
		Dim := 100 - Dim
		Manual := A_TickCount + Span
		GoSub, LoopMonitor
	}
Return
DMAuto:
	AutoDim := (AutoDim == 0) ? 1 : 0
	If (AutoDim == 0)
	{
		If (A_IsCompiled || A_ScriptName = "DimMonitor.ahk")
			Menu, Tray, Uncheck, 自動調整
		Else
			Menu, DMTray, Uncheck, 自動調整
		IniWrite, 0, % DimIni, General, AutoDim
	}
	Else
	{
		If (A_IsCompiled || A_ScriptName = "DimMonitor.ahk")
			Menu, Tray, Check, 自動調整
		Else
			Menu, DMTray, Check, 自動調整
		IniWrite, 1, % DimIni, General, AutoDim
	}
Return
DMPause:
	DMPause := (DMPause == 0) ? 1 : 0
	If (DMPause == 0)
	{
		If (A_IsCompiled || A_ScriptName = "DimMonitor.ahk")
			Menu, Tray, Uncheck, 停止
		Else
			Menu, DMTray, Uncheck, 停止
		GoTo, DMTimer
	}
	Else
	{
		If (A_IsCompiled || A_ScriptName = "DimMonitor.ahk")
			Menu, Tray, Check, 停止
		Else
			Menu, DMTray, Check, 停止
		Loop, %MonitorCount%
		{
			Gui, DimGui%A_Index%:Destroy
		}
		SetTimer, DimOnTop, OFF
	}
Return
DMExit:
	ExitApp
Return

;*** 時間帯によって明度を自動変更 ***
DimTime:
	FormatTime, Now, , Hmm
	If (AutoDim != 1)
		Return
	If (A_TickCount < Manual)
		Return
	If (Now < Day || Now >= Night)
		Dim := 100 - AutoSet			; 夜の明度
	Else
		Dim := 0
	GoSub, LoopMonitor
Return

;*** ホットキー操作 ***
#Home::							; 明度100%（不透明度0%）
	Dim = 0
	GoTo, DimHotKey
#End::							; 明度0%（不透明度100%）
	Dim = 80
	GoTo, DimHotKey
#PgDn::							; 明度を下げる（不透明度を上げる）
; ★custum mod <TOP>
	Dim += 20
; ★custum mod <END>
	If Dim > 80
		Dim = 80
	GoTo, DimHotKey
#PgUp::							; 明度を上げる（不透明度を下げる）
; ★custum mod <TOP>
	Dim -= 20
; ★custum mod <END>
	If Dim < 0
		Dim = 0
	GoTo, DimHotKey
DimHotKey:
	Manual := A_TickCount + Span
	GoSub, LoopMonitor
	AutoHideTip("明るさ：" 100 - Dim "%", 500)
Return

;全てのモニターに適用
LoopMonitor:
	Loop, %MonitorCount%
	{
		DimId := DimId%A_Index%
		WinSet, Transparent, % Dim * 255 / 100, ahk_id %DimId%
	}
Return

;*** 日出・日没時刻 ***
;日出・日入時刻取得
SunTime(Sun) {
	Pi     := ASin(1/2) * 6					;円周率
	Lat    := 35.95923333333334 * Pi /180	;緯度
	Lng    := 140.0118361111111				;経度
	D1     := 2 * Pi * (A_YDay - 81.5) / 365
	D2     := 2 * Pi * (A_YDay - 3) / 365
	K1     := -7.37 * Sin(D2)				;地球軌道が楕円による均時差
	K2     := 9.86 * Sin(2 * D1)			;地軸の傾きによる均時差
	Eot    := K1 + K2						;均時差
	Delta  := 0.4082 * Sin(D1)				;太陽赤緯
	Dt     := 1440 * (1 - ACos(Tan(Delta) * Tan(Lat)) / Pi)	;昼間の時間
	Atmspr := 0.8502 * 4 / Sqrt(1 - Sin(Lat) * Sin(Lat) / (Cos(Delta) * Cos(Delta)))	;大気補正
	Center := 720 - Eot + 4 * (135 - Lng)	;南中時刻
	Srm    := Center - Dt / 2 - Atmspr		;日入時刻の通し分
	Ssm    := Center + Dt / 2 - Atmspr		;日没時刻の通し分
	Srt    := Floor(Srm / 60) * 100 + Floor(Mod(Srm, 60))
	Sst    := Floor(Ssm / 60) * 100 + Floor(Mod(Ssm, 60))
	If (Sun = 0)
		Return Srt
	Else If (Sun = 1)
		Return Sst
}

; *** ツールチップ自動消去関数 ***
AutoHideTip(Txt, Time, X="", Y="")
{
	Tooltip, %Txt%, %X%, %Y%
	SetTimer, HideTip, -%Time%
	Return
	HideTip:
		Tooltip, 
	Return
}

DimEnd:
