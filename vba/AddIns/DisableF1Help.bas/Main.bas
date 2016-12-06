Attribute VB_Name = "Main"
Option Explicit

Const APPNAME = "F1のヘルプを無効に"

Dim Counter As Long

Public Sub InstallAddin()
    
    ModifyRegistry True
    
End Sub

Public Sub UninstallAddin()

    ModifyRegistry False
    
End Sub

Private Sub ModifyRegistry(bAdd As Boolean)

    Dim strRegPath As String
    Dim strAppPath As String
    Dim strValue As String
    Dim strText As String
    Dim strCommand As String
    
    On Error Resume Next

    strRegPath = Environ("tmp") & "\" & APPNAME & ".reg"

    strAppPath = AppPath()
    strAppPath = Replace(strAppPath, "\", "\\")
    
    ' フラグによって置換文字列を編集
    If bAdd Then
        strValue = """112,0"""
    Else
        strValue = "-"
    End If
    
    ' 出力する内容を生成
    strText = _
    "Windows Registry Editor Version 5.00" & vbCrLf & _
    "" & vbCrLf & _
    "[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\{VERSION}\Excel\DisabledShortcutKeysCheckBoxes]" & vbCrLf & _
    """{APPNAME}""={VALUE}" & vbCrLf & _
    ""
    
    strText = Replace(strText, "{VERSION}", Application.Version)
    strText = Replace(strText, "{APPNAME}", APPNAME)
    strText = Replace(strText, "{VALUE}", strValue)
    
    ' レジストリファイルを作成
    Open strRegPath For Output As #1
    Print #1, strText
    Close #1
    
    ' コマンド文字列組み立て
    strCommand = "cmd.exe /c """ & strRegPath & """"
    
    ' コマンド実行
    Shell strCommand, vbMinimizedFocus
    
End Sub

Private Function AppPath() As String

    Const EXCELFILE = "EXCEL.EXE"
    AppPath = Application.Path + "\" + EXCELFILE

End Function

Private Sub auto_open()

    On Error Resume Next
    
    SetStatusBar APPNAME + ": F1のヘルプは無効に設定されています。", "00:00:04"

End Sub

Public Sub SetStatusBar(sMsg As String, sWait As String)

    On Error Resume Next
    
    Counter = Counter + 1
    Application.StatusBar = sMsg
    Application.OnTime WakeupTime(sWait), "'ClearStatusBar'"

End Sub

Private Sub ClearStatusBar()

    On Error Resume Next
    
    Counter = Counter - 1
    If Counter = 0 Then Application.StatusBar = False

End Sub

Private Function WakeupTime(sWait)

    On Error Resume Next

    WakeupTime = TimeValue(TimeValue(Format(Now, "hh:mm:ss")) + TimeValue(sWait))

End Function


