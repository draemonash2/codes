Option Explicit

'===============================================================================
'= 設定
'===============================================================================
Const sSEND_KEY = "{F13}"
Const lSLEEP_MS = 10000
Const bOUTPUT_EXECFILE = True

'===============================================================================
'= 本処理
'===============================================================================
Const sSCRIPT_NAME = "定期キー送信"

Dim oRes
Set oRes = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery( _
        "Select * FROM Win32_Process WHERE (Caption = 'wscript.exe' OR Caption = 'cscript.exe') AND " _
      & " CommandLine LIKE '%" & WScript.ScriptName & "%'" _
    )

If oRes.Count > 1 Then
    Dim lCnt
    lCnt = 0
    Dim oProc
    For Each oProc In oRes
        lCnt = lCnt + 1
        If lCnt <> oRes.Count then
            oProc.Terminate
        End If
    Next
    Call DeleteRunningFile()
Else
    Dim objWshShell
    Set objWshShell = CreateObject("Wscript.Shell")
    Call OutputRunningFile()
    ShowToast CStr(lSLEEP_MS/1000) & "秒毎に" & sSEND_KEY & "を送信します。"
    Do
        WScript.Sleep lSLEEP_MS
        objWshShell.SendKeys(sSEND_KEY)
        ShowToast sSEND_KEY & "キーを送信しました"
    Loop
End If

ShowToast "キー送信処理を停止しました。"

'===============================================================================
'= 実行中ファイル
'===============================================================================
Const sRUN_FILE_NAME = "running_periodic_key_transmission.txt"

'実行中であることを示すファイルをデスクトップに出力する
'デスクトップが書き込み不可でもキー送信は継続させるためエラーは無視する
Sub OutputRunningFile()
    If bOUTPUT_EXECFILE = False Then
        Exit Sub
    End If

    On Error Resume Next
    Dim objTxtFile
    Set objTxtFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(GetRunningFilePath(), 2, True)
    objTxtFile.WriteLine CStr(lSLEEP_MS/1000) & "秒毎に" & sSEND_KEY & "を送信中..."
    objTxtFile.Close
    On Error Goto 0
End Sub

'実行中ファイルを削除する
Sub DeleteRunningFile()
    If bOUTPUT_EXECFILE = False Then
        Exit Sub
    End If

    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile GetRunningFilePath(), True
    On Error Goto 0
End Sub

'実行中ファイルのパスを返す
Function GetRunningFilePath()
    GetRunningFilePath = CreateObject("Wscript.Shell").SpecialFolders("Desktop") & "\" & sRUN_FILE_NAME
End Function

'===============================================================================
'= トースト通知
'===============================================================================
'通知の発行元として使用する PowerShell の AppUserModelID
Const sTOAST_APP_ID = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
'同一タグの通知は後勝ちで上書きされるため、通知センターに履歴が溜まらない
Const sTOAST_TAG = "PeriodicKeyTransmission"

'Windows のトースト通知を表示する
'WinRT の通知 API は VBScript から直接呼び出せないため PowerShell を経由する
Sub ShowToast(sMessage)
    Dim sPsCmd
    sPsCmd = "[void][Windows.UI.Notifications.ToastNotificationManager,Windows.UI.Notifications,ContentType=WindowsRuntime];" _
           & "$x=[Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02);" _
           & "$n=$x.GetElementsByTagName('text');" _
           & "[void]$n.Item(0).AppendChild($x.CreateTextNode('" & EscapePsLiteral(sSCRIPT_NAME) & "'));" _
           & "[void]$n.Item(1).AppendChild($x.CreateTextNode('" & EscapePsLiteral(sMessage) & "'));" _
           & "$t=[Windows.UI.Notifications.ToastNotification]::new($x);" _
           & "$t.Tag='" & sTOAST_TAG & "';" _
           & "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('" & sTOAST_APP_ID & "').Show($t)"

    CreateObject("Wscript.Shell").Run "powershell -NoProfile -Command """ & sPsCmd & """", 0, False
End Sub

'PowerShell のシングルクォート文字列へ埋め込める形に変換する
Function EscapePsLiteral(sText)
    EscapePsLiteral = Replace(sText, "'", "''")
End Function
