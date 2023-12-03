Option Explicit

' ==================================================================
' = 概要    管理者権限で実行する（同スクリプト内実行）
' = 引数    なし
' = 戻値                Boolean     [out]   実行結果
' = 覚書    ・同じスクリプトを管理者権限で開き直して実行する。
' =           使い勝手は良いが、利用可能な命令が足りない
' =           ソフトウェア（例:X-Finder)では使用できない。
' =         ・本関数を呼び出すスクリプト内で受け取った引数は、
' =           本関数を経由すると利用できなくなるため要注意。
' =           引数をやり取りしたい場合は、テキストファイル読み書きを利用すること。
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
' = 概要    管理者権限で実行する（別スクリプト実行）
' = 引数    sScriptPath     String      [in]    スクリプトファイルパス
' = 戻値    なし
' = 覚書    ・別のスクリプトを管理者権限で開いて実行する。
' =           主に利用可能な命令が足りないソフトウェア（例:X-Finder）
' =           にて利用することを想定している。
' =         ・本関数を呼び出すスクリプト内で受け取った引数は、
' =           本関数を経由すると利用できなくなるため要注意。
' =           引数をやり取りしたい場合は、テキストファイル読み書きを利用すること。
' = 依存    なし
' = 所属    Windows.vbs
' ==================================================================
Public Function ExecRunas2( _
    ByVal sScriptPath _
)
    Dim objShell
    Set objShell = CreateObject("Shell.Application")
    objShell.ShellExecute "wscript.exe", sScriptPath & " runas", "", "runas", 1
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
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c """ & sCommand & """")
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

' ==================================================================
' = 概要    プロセス起動確認
' = 引数    sProcessName    String      [in]    プロセス名
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Windows.vbs
' ==================================================================
Public Function ExistProcess( _
    ByVal sProcessName _
)
    Dim objService
    Dim objQfeSet
    Set objService = CreateObject("WbemScripting.SWbemLocator").ConnectServer
    Set objQfeSet = objService.ExecQuery("Select * From Win32_Process Where Caption Like '" & sProcessName & "%'")
    ExistProcess = objQfeSet.Count > 0
End Function
'   Call Test_ExistProcess()
    Private Sub Test_ExistProcess()
        MsgBox ExistProcess("wsl.exe")
    End Sub

' ==================================================================
' = 概要    WSL2 Running 待ち
' = 引数    sDistName   String  [in]    WSL2 ディストリビューション名
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Windows.vbs
' ==================================================================
Private Function WaitForWslRunning( _
    ByVal sDistName _
)
    Const sLOG_FILE_NAME = "wsl_status.log"
    
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    
    Dim sLogFilePath
    sLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sLOG_FILE_NAME
    
    Dim sStatus
    Do
        On Error Resume Next
        WScript.sleep(500)
        
        ' wslステータスコマンド リダイレクト
        objWshShell.Run "%comspec% /c wsl -l -v > """ & sLogFilePath & """", 0 , True
        
        'リダイレクトログ読み込み
        Dim adoStrm
        Set adoStrm = CreateObject("ADODB.Stream")
        adoStrm.Type = 2
        adoStrm.Charset = "UTF-16"
        adoStrm.LineSeparator = -1
        adoStrm.Open
        adoStrm.LoadFromFile sLogFilePath
        Dim sLine
        Dim sStatusLine
        Do Until adoStrm.EOS
            sLine = adoStrm.ReadText(-2)
            If InStr(sLine, "* " & sDistName) Then
                sStatusLine = sLine
            End If
        Loop
        
        'ステータス取得
        Dim sTargetStr
        sTargetStr = sLine
        Dim sSearchPattern
        sSearchPattern = "^\*\s+(" & sDistName & ")\s+(\w+)"
        oRegExp.IgnoreCase = True
        oRegExp.Global = True
        oRegExp.Pattern = sSearchPattern
        Dim oMatchResult
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        sStatus = oMatchResult(0).SubMatches(1)
        'MsgBox sStatus
        On Error Goto 0
    Loop While sStatus <> "Running"
    objFSO.DeleteFile sLogFilePath, True
End Function

