Option Explicit

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
With objWshShell.Environment("System")
                                                                                                     ' +------+------+------+------------+--------------+
                                                                                                     ' |  xf  |  ahk |  vim | codes(vbs) | updatecodes  |
                                                                                                     ' +------+------+------+------------+--------------+
    .Item("MYPATH_HIDEMARU")   = "C:\prg_exe\Hidemaru\Hidemaru.exe"                                  ' |  �Z  |  �|  |  �|  |     �|     |      �|      |
    .Item("MYPATH_WINMERGE")   = "C:\prg_exe\WinMerge\WinMergeU.exe"                                 ' |  �Z  |  �|  |  �|  |     �Z     |      �Z      |
    .Item("MYPATH_GVIM")       = "C:\prg_exe\Vim\gvim.exe"                                           ' |  �Z  |  �Z  |  �|  |     �Z     |      �|      |
    .Item("MYPATH_HNXGREP")    = "C:\prg_exe\HNXgrep\HNXgrep.exe"                                    ' |  �|  |  �|  |  �|  |     �|     |      �|      |
    .Item("MYPATH_TRESGREP")   = "C:\prg_exe\TresGrep\TresGrep.exe"                                  ' |  �Z  |  �|  |  �|  |     �|     |      �|      |
    .Item("MYPATH_EVERYTHING") = "C:\prg_exe\Everything\Everything.exe"                              ' |  �Z  |  �|  |  �|  |     �|     |      �|      |
    .Item("MYPATH_DISKINFO3")  = "C:\prg_exe\diskinfo64\DiskInfo3.exe"                               ' |  �Z  |  �|  |  �|  |     �|     |      �|      |
    .Item("MYPATH_NEEVIEW")    = "C:\prg_exe\NeeView\NeeView.exe"                                    ' |  �Z  |  �|  |  �|  |     �|     |      �|      |
    .Item("MYPATH_LINAME")     = "C:\prg_exe\LiName\LiName.exe"                                      ' |  �Z  |  �|  |  �|  |     �|     |      �|      |
    .Item("MYPATH_EXCEL")      = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"   ' |  �Z  |  �|  |  �|  |     �|     |      �|      |
    .Item("MYPATH_XF")         = "C:\prg_exe\X-Finder\XF.exe"                                        ' |  �|  |  �Z  |  �|  |     �|     |      �|      |
    .Item("MYPATH_CCALC")      = "C:\prg_exe\cCalc\cCalc.exe"                                        ' |  �|  |  �Z  |  �|  |     �|     |      �|      |
    .Item("MYPATH_RAPTURE")    = "C:\prg_exe\Rapture\rapture.exe"                                    ' |  �|  |  �Z  |  �|  |     �|     |      �|      |
    .Item("MYPATH_ITHOUGHTS")  = "C:\Program Files (x86)\toketaWare\iThoughts\iThoughts.exe"         ' |  �|  |  �Z  |  �|  |     �|     |      �|      |
    .Item("MYPATH_CTAGS")      = "C:\prg_exe\Ctags\ctags.exe"                                        ' |  �|  |  �|  |  �Z  |     �|     |      �|      |
    .Item("MYPATH_GTAGS")      = "C:\prg_exe\Gtags\bin\gtags.exe"                                    ' |  �|  |  �|  |  �Z  |     �|     |      �|      |
    .Item("MYPATH_7Z")         = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"                       ' |  �|  |  �|  |  �|  |     �Z     |      �|      |
                                                                                                     ' +------+------+------+------------+--------------+
End With

Msgbox "���ϐ���ݒ肵�܂���"

