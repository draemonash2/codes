Option Explicit

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" ) 'ExecDosCmd()

'===============================================================================
'= �{����
'===============================================================================
Dim sExecTime
Dim sBackupBatchFile
If WScript.Arguments.Count = 2 Then
    sBackupBatchFile = WScript.Arguments(0) '���ӁjBackUpFiles.bat.git_sample�̏ꍇ�́A�V���{���b�N�����N���o�R���Ȃ���΃p�X���w�肷�邱�ƁB
    sExecTime = WScript.Arguments(1)
Else
    WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
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
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
