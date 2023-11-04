Attribute VB_Name = "Mng_SysCmd"
Option Explicit

' system command library v1.02

' ==================================================================
' = �T�v    �R�}���h�����s
' = ����    sCommand    String   [in]   �R�}���h
' = ����    bGetStdout  Boolean  [in]   �W���o�͎擾�L��(�ȗ���)
' = �ߒl                String          �W���o��
' = �o��    �E��ʂ̏������s��bat�����s����ꍇ�AbGetStdout��False�ɂ��邱�ƁB
' =           �R�}���h�̎��s���ʂ��K�v�ȏꍇ�́A�R�}���h�Ƀ��_�C���N�g���܂߂邱�ƁB
' =             ��jCall ExecDosCmd("xxx.bat > xxx.log", False)
' =           �y���R�z
' =           Exec�͕W���o�͂ɂ��߂�o�b�t�@�̍ő��4096�o�C�g�ł���A
' =           ����ȏ�̃f�[�^��ǂݍ��ނ�AtEndOfStream���Ɍł܂邽�߁B
' =           https://community.cybozu.dev/t/topic/181/2
' = �ˑ�    �Ȃ�
' = ����    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmd( _
    ByVal sCommand As String, _
    Optional bGetStdOut As Boolean = True _
) As String
    If sCommand = "" Then
        ExecDosCmd = ""
    Else
        Dim sStdOutAll As String
        sStdOutAll = ""
        If bGetStdOut = True Then
            Dim oExeResult As Object
            Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c """ & sCommand & """")
            Do While Not oExeResult.StdOut.AtEndOfStream
                Dim sStdOut As String
                sStdOut = oExeResult.StdOut.ReadLine
                Debug.Print sStdOut
                sStdOutAll = sStdOutAll & vbNewLine & sStdOut
            Loop
            Set oExeResult = Nothing
        Else
            Call CreateObject("WScript.Shell").Run("%ComSpec% /c """ & sCommand & """", WaitOnReturn:=True)
        End If
        ExecDosCmd = sStdOutAll
    End If
End Function
    Private Sub Test_ExecDosCmd()
        Dim sBuf As String
        sBuf = sBuf & vbNewLine & ExecDosCmd("copy C:\Users\draem_000\Desktop\test.txt C:\Users\draem_000\Desktop\test2.txt")
        MsgBox sBuf
    End Sub

' ==================================================================
' = �T�v    �R�~�b�g�_�C�A���O��\��
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = �ˑ�    Mng_SysCmd.bas/ExecDosCmd()
' = ����    Mng_SysCmd.bas
' ==================================================================
Private Function ShowCommitDialog()
    Dim sCmdRslt As String
    Dim sCmd As String
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        sCmd = "TortoiseProc.exe " & _
               "/command:commit " & _
               "/path:""" & gtInputInfo.sTestLogDirPath & "*" & _
                            gtInputInfo.sTestDocFilePath & """ " & _
               "/closeonend:0"
               '"/logmsg:""" & "��" & """ "
    Else
        sCmd = "TortoiseProc.exe " & _
               "/command:commit " & _
               "/path:""" & gtInputInfo.sTestLogDirPath & "\" & gtInputInfo.sSubjectName & "*" & _
                            gtInputInfo.sTestDocFilePath & """ " & _
               "/closeonend:0"
               '"/logmsg:""" & "��" & """ "
    End If
    sCmdRslt = ExecDosCmd(sCmd)
End Function

