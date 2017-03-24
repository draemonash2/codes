Option Explicit

Private Function RunasCheck()
    Dim flgRunasMode
    Dim objWMI, osInfo, flag, objShell, os
    Dim strArgs
    Dim args
    
    Set args = WScript.Arguments
    
    flgRunasMode = False
    strArgs = ""
    
    ' フラグの取得
    If args.Count > 0 Then
        If UCase(args.item(0)) = "/RUNAS" Then
            flgRunasMode = True
        End If
        strArgs = strArgs & " " & args.item(0)
    End If
    
    Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set osInfo = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    flag = false
    For Each os in osInfo
        If Left(os.Version, 3) >= 6.0 Then
            flag = True
        End If
    Next
    
    Set objShell = CreateObject("Shell.Application")
    If flgRunasMode = False Then
        If flag = True Then
            objShell.ShellExecute _
            "wscript.exe", _
            """" & WScript.ScriptFullName & """" & " /RUNAS " & strArgs, "", _
            "runas", _
            1
            Wscript.Quit
        End If
    End If
End Function

'OSのバージョンを取得する
Const osWinNT = 4.0
Const osWin2k = 5.0
Const osWinXP = 5.1
Const osWin7  = 6.1
Const osWin8  = 6.2

Public Function GetOSVersion
    Dim objWMI, osInfo, os
    Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set osInfo = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each os in osInfo
        GetOSVersion = CDbl(Left(os.Version, 3))
    Next
End Function

' 管理者に昇格して実行する
Public Function ExecRunas( _
    ByVal bIsExecGui _
)
    Const cKey = "/ExecRunas"
    
    ExecRunas = False
    
    ' OS情報を取得'
    If GetOSVersion < osWin7 Then Exit Function
    
    ' 引数の処理'
    Dim sArgsStr
    sArgsStr = ""
    If WScript.Arguments.Count > 0 Then
        If WScript.Arguments.item(0) = cKey Then Exit Function  ' 実行済み'
        
        Dim i
        For i = 0 To WScript.Arguments.Count - 1
            sArgsStr = sArgsStr & " """ & WScript.Arguments.item(i) & """"
        Next
    End If
    ' Runas実行'
    If bIsExecGui = True Then
        CreateObject("Shell.Application").ShellExecute _
            "wscript.exe", _
            """" & WScript.ScriptFullName & """" & " " & cKey & " " & sArgsStr, _
            "", _
            "runas", _
            1
    Else
        CreateObject("Shell.Application").ShellExecute _
            "cscript.exe", _
            """" & WScript.ScriptFullName & """" & " " & cKey & " " & sArgsStr, _
            "", _
            "runas", _
            1
    End If
    
    ExecRunas = True
End Function

'Dos コマンド実行
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
    End Sub
