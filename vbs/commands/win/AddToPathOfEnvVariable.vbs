Option Explicit
'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\_lib\Log.vbs" )
Call Include( sMyDirPath & "\_lib\Windows.vbs" )

'==========================================================
'= �{����
'==========================================================
On Error Resume Next

'�{�X�N���v�g���Ǘ��҂Ƃ��Ď��s������
If ExecRunas( False ) Then WScript.Quit

Const ENV_NAME = "Path"         '���W�b�N��uPath�v�ȊO�͎w��ł��Ȃ����߁A�ύX�s��
Const TARGET_GROUP = "System"   'System�F���ׂẴ��[�U�[ �AUser�F���݂̃��[�U�[ �AVolatile�F���݂̃��O�I�� �AProcess�F���݂̃v���Z�X

If WScript.Arguments.Count > 1 Then
    Dim sEnvValue
    sEnvValue = WScript.Arguments(1) 'WScript.Arguments(0)�́AExecRunas()�ɂĎg�p�����
Else
    WScript.Echo "�������w�肳��Ă��܂���"
    WScript.Echo "�����𒆒f���܂�"
    WScript.Quit
End If
'MsgBox sEnvValue

'Dim oLog
'Set oLog = New LogMng
'oLog.Open "C:\Users\draem_000\Desktop\test.log", "w"

Dim objUsrEnv
If Err.Number = 0 Then
    Set objUsrEnv = WScript.CreateObject("WScript.Shell").Environment(TARGET_GROUP)
    If Err.Number = 0 Then
        Dim sEnvValOrg
        Dim sEnvValNew
        sEnvValOrg = objUsrEnv.Item(ENV_NAME)
        If InStr( sEnvValOrg, sEnvValue ) > 0 Then
            WScript.Echo ENV_NAME & "�ɑ��݂��܂�"
        Else
            sEnvValNew = sEnvValOrg & ";" & sEnvValue
        '   oLog.Puts sEnvValOrg
        '   oLog.Puts sEnvValNew
            objUsrEnv.Item(ENV_NAME) = sEnvValNew
            WScript.Echo ENV_NAME & "�ɒǉ����܂���"
        End If
    Else
        WScript.Echo "�G���[: " & Err.Description
    End If
Else
    WScript.Echo "�G���[: " & Err.Description
End If

Set objUsrEnv = Nothing

'oLog.Close
'Set oLog = Nothing

'==========================================================
'= �֐���`
'==========================================================
' �O���v���O���� �C���N���[�h�֐�
Function Include( _
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
