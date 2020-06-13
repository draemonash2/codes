Option Explicit

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
With objWshShell.Environment("System")
                                                                                                     ' +------+------+------+------------+--------------+
                                                                                                     ' |  xf  |  ahk |  vim | codes(vbs) | updatecodes  |
                                                                                                     ' +------+------+------+------------+--------------+
    .Item("MYPATH_HIDEMARU")   = "C:\prg_exe\Hidemaru\Hidemaru.exe"                                  ' |  ÅZ  |  Å|  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_WINMERGE")   = "C:\prg_exe\WinMerge\WinMergeU.exe"                                 ' |  ÅZ  |  Å|  |  Å|  |     ÅZ     |      ÅZ      |
    .Item("MYPATH_GVIM")       = "C:\prg_exe\Vim\gvim.exe"                                           ' |  ÅZ  |  ÅZ  |  Å|  |     ÅZ     |      Å|      |
    .Item("MYPATH_HNXGREP")    = "C:\prg_exe\HNXgrep\HNXgrep.exe"                                    ' |  Å|  |  Å|  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_TRESGREP")   = "C:\prg_exe\TresGrep\TresGrep.exe"                                  ' |  ÅZ  |  Å|  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_EVERYTHING") = "C:\prg_exe\Everything\Everything.exe"                              ' |  ÅZ  |  Å|  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_DISKINFO3")  = "C:\prg_exe\diskinfo64\DiskInfo3.exe"                               ' |  ÅZ  |  Å|  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_NEEVIEW")    = "C:\prg_exe\NeeView\NeeView.exe"                                    ' |  ÅZ  |  Å|  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_LINAME")     = "C:\prg_exe\LiName\LiName.exe"                                      ' |  ÅZ  |  Å|  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_EXCEL")      = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"   ' |  ÅZ  |  Å|  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_XF")         = "C:\prg_exe\X-Finder\XF.exe"                                        ' |  Å|  |  ÅZ  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_CCALC")      = "C:\prg_exe\cCalc\cCalc.exe"                                        ' |  Å|  |  ÅZ  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_RAPTURE")    = "C:\prg_exe\Rapture\rapture.exe"                                    ' |  Å|  |  ÅZ  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_ITHOUGHTS")  = "C:\Program Files (x86)\toketaWare\iThoughts\iThoughts.exe"         ' |  Å|  |  ÅZ  |  Å|  |     Å|     |      Å|      |
    .Item("MYPATH_CTAGS")      = "C:\prg_exe\Ctags\ctags.exe"                                        ' |  Å|  |  Å|  |  ÅZ  |     Å|     |      Å|      |
    .Item("MYPATH_GTAGS")      = "C:\prg_exe\Gtags\bin\gtags.exe"                                    ' |  Å|  |  Å|  |  ÅZ  |     Å|     |      Å|      |
    .Item("MYPATH_7Z")         = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"                       ' |  Å|  |  Å|  |  Å|  |     ÅZ     |      Å|      |
                                                                                                     ' +------+------+------+------------+--------------+
End With

Msgbox "ä¬ã´ïœêîÇê›íËÇµÇ‹ÇµÇΩ"

