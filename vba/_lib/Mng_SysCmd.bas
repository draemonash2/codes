Attribute VB_Name = "Mng_SysCmd"
Option Explicit

' system command library v1.03

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'for ExecDosCmdRunas()
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long 'for ExecDosCmdRunas()

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
' = �T�v    �R�}���h�����s�i�Ǘ��Ҍ����j
' = ����    asCommands()    String   [in] ���s�R�}���h
' = ����    bDelFiles       Boolean  [in] Bat/Log�t�@�C���폜(�ȗ���)
' = �ߒl                    String        �W���o�́��W���G���[�o��
' = �o��    �EDesktop�t�H���_�p�X�ɋ󔒂��܂܂��ꍇ�́A���삵�Ȃ��B
' = �ˑ�    �Ȃ�
' = �ˑ�    Mng_FileSys.bas/OutputTxtFile()
' = ����    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmdRunas( _
    ByRef asCommands() As String, _
    Optional bDelFiles As Boolean = True _
) As String
    Const sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME As String = "CmdExeBatRunas"
    If Sgn(asCommands) = 0 Then
        ExecDosCmdRunas = ""
    Else
        If UBound(asCommands) < 0 Then
            ExecDosCmdRunas = ""
        Else
            Dim objWshShell
            Set objWshShell = CreateObject("WScript.Shell")
            Dim objFSO
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            
            Dim sBatFilePath As String
            Dim sLogFilePath As String
            sBatFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME & ".bat"
            sLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME & ".log"
            
            '�u@echo off�v�}��
            ReDim Preserve asCommands(UBound(asCommands) + 1)
            Dim lIdx As Long
            For lIdx = UBound(asCommands) To (LBound(asCommands) + 1) Step -1
                asCommands(lIdx) = asCommands(lIdx - 1)
            Next lIdx
            asCommands(0) = "@echo off"
            
            'BAT�t�@�C���쐬
            Call OutputTxtFile(sBatFilePath, asCommands)
            Do While Not objFSO.FileExists(sBatFilePath)
                Sleep 100
            Loop
            
            'BAT�t�@�C�����s
            ShellExecute 0, "runas", sBatFilePath, " > " & sLogFilePath & " 2>&1", vbNullString, 1
            
            'LOG�t�@�C���o�͑҂�
            Do While Not objFSO.FileExists(sLogFilePath)
                Sleep 100
            Loop
            
            'LOG�t�@�C���Ǎ���
            Dim objTxtFile
            Set objTxtFile = objFSO.OpenTextFile(sLogFilePath, 1, True)
            Dim sStdOutAll As String
            sStdOutAll = ""
            Dim sLine As String
            Do Until objTxtFile.AtEndOfStream
                sLine = objTxtFile.ReadLine
                'MsgBox sLine
                If sStdOutAll = "" Then
                    sStdOutAll = sLine
                Else
                    sStdOutAll = sStdOutAll & vbNewLine & sLine
                End If
            Loop
            'MsgBox sStdOutAll
            objTxtFile.Close
            
            'BAT�t�@�C��/LOG�t�@�C���폜
            If bDelFiles = True Then
                Kill sBatFilePath
                Kill sLogFilePath
            End If
            
            ExecDosCmdRunas = sStdOutAll
        End If
    End If
End Function
    Private Sub Test_ExecDosCmdRunas()
        Dim asCommands() As String
        
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(0)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target.txt"""
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(1)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target.txt"""
        asCommands(1) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source2.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target2.txt"""
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(1)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target.txt"""
        asCommands(1) = "mklink ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\source2.txt"" ""C:\Users\draem\OneDrive\�f�X�N�g�b�v\target2.txt"""
        MsgBox ExecDosCmdRunas(asCommands, False)
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

