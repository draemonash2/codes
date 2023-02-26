; DimMonitor.ahk 2.1.0
; https://sites.google.com/site/bucuerider/

#NoEnv
#SingleInstance force

Dim = 0
Span=2			;手動調整後に自動調整を停止する間隔（時間単位）
DimColor=000000	;フィルタの色（HTMLカラーコード参照）

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
	;	MsgBox %DimId%
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

; *** ツールチップ自動消去関数 ***
AutoHideTip(Txt, Time, X="", Y="")
{
	Tooltip, %Txt%, %X%, %Y%
	SetTimer, HideTip, -%Time%
	Return
}

HideTip:
	Tooltip, 
	Return

DimEnd:
