Attribute VB_Name = "Mng_SysCmd"
Option Explicit

' system command library v1.01

' ==================================================================
' = �T�v    �R�}���h�����s
' = ����    sCommand    String   [in]   �R�}���h
' = �ߒl                String          �W���o��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmd( _
    ByVal sCommand As String _
) As String
    Dim oExeResult As Object
    Dim sStrOut As String
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
      sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
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

