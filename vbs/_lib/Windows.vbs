Option Explicit

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

' ==================================================================
' = 概要    Dos コマンド実行
' = 引数    なし
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Windows.vbs
' ==================================================================
Public Function ExecDosCmd( _
    ByVal sCommand _
)
    Dim oExeResult
    Dim sStrOut
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
        sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function
'   Call Test_ExecDosCmd()
    Private Sub Test_ExecDosCmd()
        Msgbox ExecDosCmd( "copy ""C:\Users\draem_000\Desktop\test.txt"" ""C:\Users\draem_000\Desktop\test2.txt""" )
        'Msgbox ExecDosCmd( "C:\codes\vbs\_lib\test.bat" )
    End Sub
