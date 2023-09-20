Option Explicit

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" ) 'ExecDosCmd()

'===============================================================================
'= 本処理
'===============================================================================
Dim sExecTime
Dim sBackupBatchFile
If WScript.Arguments.Count = 2 Then
    sBackupBatchFile = WScript.Arguments(0) '注意）BackUpFiles.bat.git_sampleの場合は、シンボリックリンクを経由しない絶対パスを指定すること。
    sExecTime = WScript.Arguments(1)
Else
    WScript.Echo "引数を指定してください。プログラムを中断します。"
    WScript.Quit
End If
'MsgBox sBackupBatchFile & vbNewLine & sExecTime

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

Do While True
    Dim sCurTime
    Dim sTrgtDateTime
    sCurTime = Now()
    sTrgtDateTime = Left(sCurTime, InStr(sCurTime, " ")) & " " & sExecTime
    'MsgBox sTrgtDateTime
    Dim lDateDiff
    lDateDiff = DateDiff("n", sCurTime, sTrgtDateTime)
    'MsgBox lDateDiff
    If lDateDiff = 0 Then
        Dim sCmd
        sCmd = """" & sBackupBatchFile & """ ""Scheduled backup."""
        'MsgBox "The time has come!" & vbNewLine & sCmd
        Call ExecDosCmd(sCmd)
    End If
    WScript.sleep(60000) '60[s]
Loop

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
