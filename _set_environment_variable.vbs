Option Explicit

Call ExecRunas()

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
With objWshShell.Environment("System")
                                                                                                     ' +------+------+------+------------+--------------+
                                                                                                     ' |  xf  |  ahk |  vim | codes(vbs) | updatecodes  |
                                                                                                     ' +------+------+------+------------+--------------+
    .Item("MYPATH_HIDEMARU")   = "C:\prg_exe\Hidemaru\Hidemaru.exe"                                  ' |  Z  |  |  |  |  |     |     |      |      |
    .Item("MYPATH_WINMERGE")   = "C:\prg_exe\WinMerge\WinMergeU.exe"                                 ' |  Z  |  |  |  |  |     Z     |      Z      |
    .Item("MYPATH_GVIM")       = "C:\prg_exe\Vim\gvim.exe"                                           ' |  Z  |  Z  |  |  |     Z     |      |      |
    .Item("MYPATH_TRESGREP")   = "C:\prg_exe\TresGrep\TresGrep.exe"                                  ' |  Z  |  |  |  |  |     |     |      |      |
    .Item("MYPATH_EVERYTHING") = "C:\prg_exe\Everything\Everything.exe"                              ' |  Z  |  |  |  |  |     |     |      |      |
    .Item("MYPATH_DISKINFO3")  = "C:\prg_exe\diskinfo64\DiskInfo3.exe"                               ' |  Z  |  |  |  |  |     |     |      |      |
    .Item("MYPATH_NEEVIEW")    = "C:\prg_exe\NeeView\NeeView.exe"                                    ' |  Z  |  |  |  |  |     |     |      |      |
    .Item("MYPATH_MASSIGRA")   = "C:\prg_exe\MassiGra\MassiGra.exe"                                  ' |  Z  |  |  |  |  |     |     |      |      |
    .Item("MYPATH_LINAME")     = "C:\prg_exe\LiName\LiName.exe"                                      ' |  Z  |  |  |  |  |     |     |      |      |
    .Item("MYPATH_EXCEL")      = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"   ' |  Z  |  |  |  |  |     |     |      |      |
    .Item("MYPATH_XF")         = "C:\prg_exe\X-Finder\XF.exe"                                        ' |  |  |  Z  |  |  |     |     |      |      |
    .Item("MYPATH_CCALC")      = "C:\prg_exe\cCalc\cCalc.exe"                                        ' |  |  |  Z  |  |  |     |     |      |      |
    .Item("MYPATH_RAPTURE")    = "C:\prg_exe\Rapture\rapture.exe"                                    ' |  |  |  Z  |  |  |     |     |      |      |
    .Item("MYPATH_ITHOUGHTS")  = "C:\prg_exe\iThoughts\iThoughts.exe"                                ' |  |  |  Z  |  |  |     |     |      |      |
    .Item("MYPATH_CTAGS")      = "C:\prg_exe\Ctags\ctags.exe"                                        ' |  |  |  |  |  Z  |     |     |      |      |
    .Item("MYPATH_GTAGS")      = "C:\prg_exe\Gtags\bin\gtags.exe"                                    ' |  |  |  |  |  Z  |     |     |      |      |
    .Item("MYPATH_7Z")         = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"                       ' |  |  |  |  |  |  |     Z     |      |      |
                                                                                                     ' +------+------+------+------------+--------------+
End With

Msgbox "ŠÂ‹«•Ï”‚ðÝ’è‚µ‚Ü‚µ‚½"

' ==================================================================
' = ŠT—v    ŠÇ—ŽÒŒ ŒÀ‚ÅŽÀs‚·‚é
' = ˆø”    ‚È‚µ
' = –ß’l    ‚È‚µ
' = –ß’l                Boolean     [out]   ŽÀsŒ‹‰Ê
' = Šo‘    Ž©“®“I‚Éˆø”‚É‰e‹¿‚ð‹y‚Ú‚·‚½‚ßA—v’ˆÓ
' = ˆË‘¶    ‚È‚µ
' = Š‘®    Windows.vbs
' ==================================================================
Public Function ExecRunas()
    Dim oArgs
    Dim bIsRunas
    Dim sArgs
    
    bIsRunas = False
    sArgs = ""
    Set oArgs = WScript.Arguments
    
    ' ƒtƒ‰ƒO‚ÌŽæ“¾
    If oArgs.Count > 0 Then
        If UCase(oArgs.item(0)) = "/RUNAS" Then
            bIsRunas = True
        End If
        sArgs = sArgs & " " & oArgs.item(0)
    End If
    
    Dim bIsExecutableOs
    bIsExecutableOs = false
    Dim oOsInfos
    Set oOsInfos = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_OperatingSystem")
    Dim oOs
    For Each oOs in oOsInfos
        If Left(oOs.Version, 3) >= 6.0 Then
            bIsExecutableOs = True
        End If
    Next
    
    Dim oWshShell
    Set oWshShell = CreateObject("Shell.Application")
    ExecRunas = False
    If bIsRunas = False Then
        If bIsExecutableOs = True Then
            oWshShell.ShellExecute _
            "wscript.exe", _
            """" & WScript.ScriptFullName & """" & " /RUNAS " & sArgs, "", _
            "runas", _
            1
            ExecRunas = True
            Wscript.Quit
        End If
    End If
End Function
