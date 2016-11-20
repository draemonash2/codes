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

Function GetOSVersion
    Dim objWMI, osInfo, os
    Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set osInfo = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each os in osInfo
        GetOSVersion = CDbl(Left(os.Version, 3))
    Next
End Function

' 管理者に昇格して実行する
Function ExecRunas
    Const cKey = "/ExecRunas"
    Dim s
    
    ExecRunas = False
    
    ' OS情報を取得'
    If GetOSVersion < osWin7 Then Exit Function
    
    ' 引数の処理'
    s = ""
    If WScript.Arguments.Count > 0 Then
        If WScript.Arguments.item(0) = cKey Then Exit Function  ' 実行済み'
        
        Dim i
        For i = 0 To WScript.Arguments.Count -1
            s = s & " """ & WScript.Arguments.item(i) & """"
        Next
    End If
    ' Runas実行'
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & " " &cKey& " " & s, "", "runas", 1
    
    ExecRunas = True
End Function

