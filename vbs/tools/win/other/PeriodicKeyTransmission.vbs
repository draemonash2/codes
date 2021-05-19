Option Explicit

'===============================================================================
'= 設定
'===============================================================================
Const sSEND_KEY = "{F13}"
Const lSLEEP_MS = 10000

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
Else
    Dim vAnswer
    vAnswer = MsgBox(CStr(lSLEEP_MS/1000) & "秒毎に" & sSEND_KEY &"を送信します。" , vbYesNo + vbQuestion, sSCRIPT_NAME)
    If vAnswer = vbYes Then
        Dim objWshShell
        Set objWshShell = CreateObject("Wscript.Shell")
        Do
            WScript.Sleep lSLEEP_MS
            objWshShell.SendKeys(sSEND_KEY)
            objWshShell.Popup sSEND_KEY & "キーを送信しました", 3, sSCRIPT_NAME, vbInformation
        Loop
    End If
End If

MsgBox "キー送信処理を停止しました。", vbOkOnly, sSCRIPT_NAME

