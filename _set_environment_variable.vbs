Option Explicit

Call ExecRunas()

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
With objWshShell.Environment("System")
                                                                                                     ' +------+------+------+------------+--------------+
                                                                                                     ' |  xf  |  ahk |  vim | codes(vbs) | updatecodes  |
                                                                                                     ' +------+------+------+------------+--------------+
    .Item("MYPATH_HIDEMARU")   = "C:\prg_exe\Hidemaru\Hidemaru.exe"                                  ' |  〇  |  −  |  −  |     −     |      −      |
    .Item("MYPATH_WINMERGE")   = "C:\prg_exe\WinMerge\WinMergeU.exe"                                 ' |  〇  |  −  |  −  |     〇     |      〇      |
    .Item("MYPATH_GVIM")       = "C:\prg_exe\Vim\gvim.exe"                                           ' |  〇  |  〇  |  −  |     〇     |      −      |
    .Item("MYPATH_TRESGREP")   = "C:\prg_exe\TresGrep\TresGrep.exe"                                  ' |  〇  |  −  |  −  |     −     |      −      |
    .Item("MYPATH_EVERYTHING") = "C:\prg_exe\Everything\Everything.exe"                              ' |  〇  |  −  |  −  |     −     |      −      |
    .Item("MYPATH_DISKINFO3")  = "C:\prg_exe\diskinfo64\DiskInfo3.exe"                               ' |  〇  |  −  |  −  |     −     |      −      |
    .Item("MYPATH_NEEVIEW")    = "C:\prg_exe\NeeView\NeeView.exe"                                    ' |  〇  |  −  |  −  |     −     |      −      |
    .Item("MYPATH_MASSIGRA")   = "C:\prg_exe\MassiGra\MassiGra.exe"                                  ' |  〇  |  −  |  −  |     −     |      −      |
    .Item("MYPATH_LINAME")     = "C:\prg_exe\LiName\LiName.exe"                                      ' |  〇  |  −  |  −  |     −     |      −      |
    .Item("MYPATH_EXCEL")      = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"   ' |  〇  |  −  |  −  |     −     |      −      |
    .Item("MYPATH_XF")         = "C:\prg_exe\X-Finder\XF.exe"                                        ' |  −  |  〇  |  −  |     −     |      −      |
    .Item("MYPATH_CCALC")      = "C:\prg_exe\cCalc\cCalc.exe"                                        ' |  −  |  〇  |  −  |     −     |      −      |
    .Item("MYPATH_RAPTURE")    = "C:\prg_exe\Rapture\rapture.exe"                                    ' |  −  |  〇  |  −  |     −     |      −      |
    .Item("MYPATH_ITHOUGHTS")  = "C:\prg_exe\iThoughts\iThoughts.exe"                                ' |  −  |  〇  |  −  |     −     |      −      |
    .Item("MYPATH_CTAGS")      = "C:\prg_exe\Ctags\ctags.exe"                                        ' |  −  |  −  |  〇  |     −     |      −      |
    .Item("MYPATH_GTAGS")      = "C:\prg_exe\Gtags\bin\gtags.exe"                                    ' |  −  |  −  |  〇  |     −     |      −      |
    .Item("MYPATH_7Z")         = "C:\prg_exe\7-ZipPortable\App\7-Zip64\7z.exe"                       ' |  −  |  −  |  −  |     〇     |      −      |
                                                                                                     ' +------+------+------+------------+--------------+
End With

Msgbox "環境変数を設定しました"

' ==================================================================
' = 概要    管理者権限で実行する
' = 引数    なし
' = 戻値    なし
' = 戻値                Boolean     [out]   実行結果
' = 覚書    自動的に引数に影響を及ぼすため、要注意
' = 依存    なし
' = 所属    Windows.vbs
' ==================================================================
Public Function ExecRunas()
    Dim oArgs
    Dim bIsRunas
    Dim sArgs
    
    bIsRunas = False
    sArgs = ""
    Set oArgs = WScript.Arguments
    
    ' フラグの取得
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
