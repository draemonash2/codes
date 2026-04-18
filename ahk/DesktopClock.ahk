; [Help] https://ahkscript.github.io/ja/docs/v2/

;	#NoTrayIcon						; スクリプトのタスクトレイアイコンを非表示にする。
	#Warn All						; Enable warnings to assist with detecting common errors.
	#SingleInstance force			; このスクリプトが再度呼び出されたらリロードして置き換え
	#WinActivateForce				; ウィンドウのアクティブ化時に、穏やかな方法を試みるのを省略して常に強制的な方法でアクティブ化を行う。（タスクバーアイコンが点滅する現象が起こらなくなる）
	SendMode "Input"				; WindowsAPIの SendInput関数を利用してシステムに一連の操作イベントをまとめて送り込む方式。
	TraySetIcon "DesktopClock.ico"

^+!Enter::ReloadMe()
ReloadMe() ; {{{
{
	Reload
	Sleep 1000 ; リロードに成功した場合、リロードはスリープ中にこのインスタンスを閉じるので、以下の行に到達することはない
	MsgBox "スクリプト" . A_ScriptName . "の再読み込みに失敗しました"
} ; }}}

global gsDESKTOPCLOCK_INFO := [
	; ClockGui(x, y, fontSize, width[, height])
	;   height=0 (default): auto-sized to font; height>0: explicit window height
;	ClockGui(1732, 742, 33, 200),		; Main
;	ClockGui(851, 2350, 30, 200),		; Mobile
;	ClockGui(3025, -449, 50, 300),		; DualUp
	ClockGui(4539, -612, 90, 450),		; 4K
]

global gsDESKTOPDATE_INFO := [
	; DateGui(x, y, fontSize, width[, height])
	;   height=0 (default): auto-sized to font; height>0: explicit window height
;	DateGui(1732, 700, 20, 200),		; Main
;	DateGui(851, 2320, 18, 200),		; Mobile
;	DateGui(3025, -490, 30, 300),		; DualUp
	DateGui(4479, -680, 50, 450),		; 4K
]

StartDesktopClock()

; デスクトップ時計
class ClockGui {
	__New(iX, iY, iFontSize := 60, iWidth := 300, iHeight := 0) {
		this.gui := Gui("+AlwaysOnTop -Caption +ToolWindow")
		this.gui.BackColor := "Black"
		this.gui.MarginX := 0
		this.gui.MarginY := 0
		this.gui.SetFont("s" iFontSize, "Segoe UI")
	;	this.gui.SetFont("s" iFontSize, "DSEG7 Classic-Bold")
		this.clockText := this.gui.AddText("cWhite Center w" iWidth, "")
		; Make black background fully transparent — only the white text is visible
		WinSetTransColor("Black 220", this.gui.Hwnd)
		; iHeight=0: auto-size; iHeight>0: set explicit window height
		sShowOpt := "x" iX " y" iY
		if (iHeight > 0)
			sShowOpt .= " h" iHeight
		this.gui.Show(sShowOpt)
		this.Update()
	}
	Update() {
		this.clockText.Text := FormatTime(, "HH:mm:ss")
	}
	CheckMouseOver() {
		MouseGetPos(&mx, &my)
		WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " this.gui.Hwnd)
		if (mx >= wx && mx < wx + ww && my >= wy && my < wy + wh)
			WinSetTransColor("Black 30", this.gui.Hwnd)   ; fade text on hover
		else
			WinSetTransColor("Black 220", this.gui.Hwnd)  ; restore text opacity
	}
	IsHwnd(hwnd) {
		return hwnd = this.gui.Hwnd
	}
}
; Desktop date display (MM/DD (ddd) format)
class DateGui {
	__New(iX, iY, iFontSize := 40, iWidth := 300, iHeight := 0) {
		this.gui := Gui("+AlwaysOnTop -Caption +ToolWindow")
		this.gui.BackColor := "Black"
		this.gui.MarginX := 0
		this.gui.MarginY := 0
		this.gui.SetFont("s" iFontSize, "Segoe UI")
		this.dateText := this.gui.AddText("cWhite Center w" iWidth, "")
		; Make black background fully transparent — only the white text is visible
		WinSetTransColor("Black 220", this.gui.Hwnd)
		; iHeight=0: auto-size; iHeight>0: set explicit window height
		sShowOpt := "x" iX " y" iY
		if (iHeight > 0)
			sShowOpt .= " h" iHeight
		this.gui.Show(sShowOpt)
		this.Update()
	}
	Update() {
		static dayNames := ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
		this.dateText.Text := FormatTime(, "MM/dd") " (" dayNames[A_WDay] ")"
	}
	CheckMouseOver() {
		MouseGetPos(&mx, &my)
		WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " this.gui.Hwnd)
		if (mx >= wx && mx < wx + ww && my >= wy && my < wy + wh)
			WinSetTransColor("Black 30", this.gui.Hwnd)   ; fade text on hover
		else
			WinSetTransColor("Black 220", this.gui.Hwnd)  ; restore text opacity
	}
	IsHwnd(hwnd) {
		return hwnd = this.gui.Hwnd
	}
}
StartDesktopClock() {
	OnMessage(0x0084, WM_NCHITTEST)  ; WM_NCHITTEST
	SetTimer(_CheckMouseOverAllClocks, 100)
	SetTimer(_UpdateAllClocks, 1000)
}
; Make gsDESKTOPCLOCK_INFO / gsDESKTOPDATE_INFO transparent on mouse hover (polling-based)
_CheckMouseOverAllClocks() {
	global gsDESKTOPCLOCK_INFO, gsDESKTOPDATE_INFO
	for clock in gsDESKTOPCLOCK_INFO
		clock.CheckMouseOver()
	for date in gsDESKTOPDATE_INFO
		date.CheckMouseOver()
}
; Update all gsDESKTOPCLOCK_INFO / gsDESKTOPDATE_INFO every second
_UpdateAllClocks() {
	global gsDESKTOPCLOCK_INFO, gsDESKTOPDATE_INFO
	for clock in gsDESKTOPCLOCK_INFO
		clock.Update()
	for date in gsDESKTOPDATE_INFO
		date.Update()
}
; Allow dragging each clock/date window with Ctrl+drag
WM_NCHITTEST(wParam, lParam, msg, hwnd) {
	global gsDESKTOPCLOCK_INFO, gsDESKTOPDATE_INFO
	for clock in gsDESKTOPCLOCK_INFO {
		if clock.IsHwnd(hwnd) && GetKeyState("Ctrl")
			return 2  ; HTCAPTION - treat entire window as title bar to enable dragging
	}
	for date in gsDESKTOPDATE_INFO {
		if date.IsHwnd(hwnd) && GetKeyState("Ctrl")
			return 2  ; HTCAPTION - treat entire window as title bar to enable dragging
	}
}
