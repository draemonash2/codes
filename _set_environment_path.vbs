Option Explicit

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
With objWshShell.Environment("System")
	.Item("MYSYSPATH_HIDEMARU") =	"C:\prg_exe\Hidemaru\Hidemaru.exe"
	.Item("MYSYSPATH_WINMERGE") =	"C:\prg_exe\WinMerge\WinMergeU.exe"
	.Item("MYSYSPATH_GVIM") =		"C:\prg_exe\Vim\gvim.exe"
	.Item("MYSYSPATH_HNXGREP") =	"C:\prg_exe\HNXgrep\HNXgrep.exe"
	.Item("MYSYSPATH_TRESGREP") =	"C:\prg_exe\TresGrep\TresGrep.exe"
	.Item("MYSYSPATH_EVERYTHING") =	"C:\prg_exe\Everything\Everything.exe"
	.Item("MYSYSPATH_DISKINFO3") =	"C:\prg_exe\diskinfo64\DiskInfo3.exe"
	.Item("MYSYSPATH_7Z") =			"C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"
	.Item("MYSYSPATH_NEEVIEW") =	"C:\prg_exe\NeeView\NeeView.exe"
	.Item("MYSYSPATH_LINAME") =		"C:\prg_exe\LiName\LiName.exe"
	.Item("MYSYSPATH_CTAGS") =		"C:\prg_exe\Ctags\ctags.exe"
	.Item("MYSYSPATH_GTAGS") =		"C:\prg_exe\Gtags\bin\gtags.exe"
	.Item("MYSYSPATH_XF") =			"C:\prg_exe\X-Finder\XF.exe"
	.Item("MYSYSPATH_CCALC") =		"C:\prg_exe\cCalc\cCalc.exe"
	.Item("MYSYSPATH_RAPTURE") =	"C:\prg_exe\Rapture\rapture.exe"
	.Item("MYSYSPATH_EXCEL") =		"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
	.Item("MYSYSPATH_ITHOUGHTS") =	"C:\Program Files (x86)\toketaWare\iThoughts\iThoughts.exe"
End With

Msgbox "ä¬ã´ïœêîÇê›íËÇµÇ‹ÇµÇΩ"


