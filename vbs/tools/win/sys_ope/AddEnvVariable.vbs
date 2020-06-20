Option Explicit

'usage
' cscript.exe .\AddEnvVariable.vbs <env_value> [<env_name>] [<target_group>]
' 
'usage ex.
' cscript.exe .\AddEnvVariable.vbs c:\codes
' cscript.exe .\AddEnvVariable.vbs c:\codes CODES
' cscript.exe .\AddEnvVariable.vbs c:\codes CODES System

'==========================================================
'= �ݒ�
'==========================================================
Const DEFAULT_ENV_NAME = "Path"
Const DEFAULT_TARGET_GROUP = "User"   'System�F���ׂẴ��[�U�[ �AUser�F���݂̃��[�U�[ �AVolatile�F���݂̃��O�I�� �AProcess�F���݂̃v���Z�X

'==========================================================
'= �C���N���[�h
'==========================================================
'Call Include( "C:\codes\vbs\_lib\Log.vbs" )        'Class LogMng

'==========================================================
'= �{����
'==========================================================
On Error Resume Next

Dim sEnvValue
Dim sEnvName
Dim sTargetGroup
sEnvName = DEFAULT_ENV_NAME
sTargetGroup = DEFAULT_TARGET_GROUP
If WScript.Arguments.Count = 1 Then
    sEnvValue = WScript.Arguments(0)
ElseIf WScript.Arguments.Count = 2 Then
    sEnvValue = WScript.Arguments(0)
    sEnvName = WScript.Arguments(1)
ElseIf WScript.Arguments.Count > 3 Then
    sEnvValue = WScript.Arguments(0)
    sEnvName = WScript.Arguments(1)
    sTargetGroup = WScript.Arguments(2)
Else
    WScript.StdOut.WriteLine "�������w�肳��Ă��܂���"
    WScript.StdOut.WriteLine "�����𒆒f���܂�"
    WScript.Quit
End If
'MsgBox sEnvValue

'Dim oLog
'Set oLog = New LogMng
'oLog.Open "C:\Users\draem_000\Desktop\test.log", "w"

Dim objUsrEnv
Set objUsrEnv = WScript.CreateObject("WScript.Shell").Environment(sTargetGroup)
If Err.Number = 0 Then
    'Do Nothing
Else
    WScript.StdOut.WriteLine "�G���[: " & Err.Description
    WScript.StdOut.WriteLine "���ϐ��G���["
    WScript.Quit
End If

Dim sEnvValOrg
Dim sEnvValNew
sEnvValOrg = objUsrEnv.Item(sEnvName)
If sEnvValOrg = "" Then
    objUsrEnv.Item(sEnvName) = sEnvValue
    WScript.StdOut.WriteLine sEnvName & "��ǉ����܂���"
Else
    If InStr( sEnvValOrg, sEnvValue ) > 0 Then
        WScript.StdOut.WriteLine sEnvValue & "��" & sEnvName & "�ɑ��݂��܂�"
    Else
        sEnvValNew = sEnvValOrg & ";" & sEnvValue
    '   oLog.Puts sEnvValOrg
    '   oLog.Puts sEnvValNew
        objUsrEnv.Item(sEnvName) = sEnvValNew
        WScript.StdOut.WriteLine sEnvValue & "��" & sEnvName & "�ɒǉ����܂���"
    End If
End If

Set objUsrEnv = Nothing

'oLog.Close
'Set oLog = Nothing

'==========================================================
'= �C���N���[�h�֐�
'==========================================================
Private Function Include( _
    ByVal sOpenFile _
)
    Dim objFSO
    Dim objVbsFile
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function
