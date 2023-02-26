SendMode "Input"					; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。
#SingleInstance force

Dim := 0
DimId := 0

;*** フィルタ生成 ***
DimMon_GenFilter()

;*** ホットキー操作 ***
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

;*** 内部関数 ***
DimMon_GenFilter()
{
	global MonitorCount := MonitorGetCount()
	aDimGui := Array()
	global aDimId := Array()
;	MsgBox "MonitorCount = " . MonitorCount
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

