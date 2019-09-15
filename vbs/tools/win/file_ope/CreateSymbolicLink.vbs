'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################
Const OBJECT_SUFFIX = " - �V���{���b�N�����N"

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�V���{���b�N�����N�쐬"

'*** �Ǘ��҂Ƃ��Ď��s ***
Call ExecRunas()

'*** �t�@�C��/�t�H���_���擾 ***
DIm cFilePaths
If EXECUTION_MODE = 0 Then 'Explorer������s
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    Dim sArg
    For Each sArg In WScript.Arguments
        If sArg = "/RUNAS" Then
            'Do Nothing
        Else
            cFilePaths.add sArg
        End If
    Next
ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
    Set cFilePaths = WScript.Col( WScript.Env("Selected") )
Else '�f�o�b�O���s
    MsgBox "�f�o�b�O���[�h�ł��B"
    Set cFilePaths = CreateObject("System.Collections.ArrayList")
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Dim sDesktop
    sDesktop = objWshShell.SpecialFolders("Desktop")
    objWshShell.Run "cmd /c echo.> """ & sDesktop & "\test.txt""", 0, True
    objWshShell.Run "cmd /c mkdir """ & sDesktop & "\test2""", 0, True
    cFilePaths.Add sDesktop & "\test.txt"
    cFilePaths.Add sDesktop & "\test2"
End If
'������debug������
'For Each sArg In cFilePaths
'    msgbox sArg
'Next
'������debug������

'*** �t�@�C���p�X�`�F�b�N ***
If cFilePaths.Count = 0 Then
    MsgBox "�I�u�W�F�N�g���I������Ă��܂���", vbYes, PROG_NAME
    MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
    WScript.Quit
End If

'*** �V���{���b�N�����N�쐬 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim oObjPath
Dim lObjType '0:notexists 1:file 2:folder
For Each oObjPath In cFilePaths
    lObjType = GetFileOrFolder( oObjPath )
    
    Dim sDstPath
    Dim sSrcPath
    sDstPath = oObjPath
    Dim sCmd
    If lObjType = 1 Then 'file
        sSrcPath = objFSO.GetParentFolderName( oObjPath ) & "\" & _
                   objFSO.GetBaseName( oObjPath ) & OBJECT_SUFFIX & "." & _
                   objFSO.GetExtensionName( oObjPath )
        sCmd = "mklink """ & sSrcPath & """ """ & sDstPath & """"
    ElseIf lObjType = 2 Then 'folder
        sSrcPath = oObjPath & OBJECT_SUFFIX
        sCmd = "mklink /d """ & sSrcPath & """ """ & sDstPath & """"
    Else 'not exists
        MsgBox "�I�u�W�F�N�g�����݂��܂���", vbYes, PROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
        WScript.Quit
    End If
    '������debug������
    'msgbox sCmd
    '������debug������
    call ExecDosCmd( sCmd )
Next

MsgBox "�V���{���b�N�����N���쐬���܂���", vbYes, PROG_NAME

'####################################################################
'### �֐���`
'####################################################################
' ==================================================================
' = �T�v    �Ǘ��Ҍ����Ŏ��s����
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �ߒl                Boolean     [out]   ���s����
' = �o��    �����I�Ɉ����ɉe�����y�ڂ����߁A�v����
' = �ˑ�    �Ȃ�
' = ����    Windows.vbs
' ==================================================================
Public Function ExecRunas()
    Dim oArgs
    Dim bIsRunas
    Dim sArgs
    
    bIsRunas = False
    sArgs = ""
    Set oArgs = WScript.Arguments
    
    ' �t���O�̎擾
    If oArgs.Count > 0 Then
        If UCase(oArgs.item(0)) = "/RUNAS" Then
            bIsRunas = True
        End If
        sArgs = sArgs & " " & oArgs.item(0)
    End If
    
    Dim bIsExecutableOs
    bIsExecutableOs = false
    Dim oOsInfos
    Set oOsInfos = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_OperatingSystem")
    Dim oOs
    For Each oOs in oOsInfos
        If Left(oOs.Version, 3) >= 6.0 Then
            bIsExecutableOs = True
        End If
    Next
    
    Dim oWshShell
    Set oWshShell = CreateObject("Shell.Application")
    ExecRunas = False
    If bIsRunas = False Then
        If bIsExecutableOs = True Then
            oWshShell.ShellExecute _
            "wscript.exe", _
            """" & WScript.ScriptFullName & """" & " /RUNAS " & sArgs, "", _
            "runas", _
            1
            ExecRunas = True
            Wscript.Quit
        End If
    End If
End Function

' ==================================================================
' = �T�v    Dos �R�}���h���s
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Windows.vbs
' ==================================================================
Public Function ExecDosCmd( _
    ByVal sCommand _
)
    Dim oExeResult
    Dim sStrOut
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
        sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function
'   Call Test_ExecDosCmd()
    Private Sub Test_ExecDosCmd()
        Msgbox ExecDosCmd( "copy ""C:\Users\draem_000\Desktop\test.txt"" ""C:\Users\draem_000\Desktop\test2.txt""" )
        'Msgbox ExecDosCmd( "C:\codes\vbs\_lib\test.bat" )
    End Sub

' ==================================================================
' = �T�v    �t�@�C�����t�H���_���𔻒肷��
' = ����    sChkTrgtPath    String      [in]    �`�F�b�N�Ώۃt�H���_
' = �ߒl                    Long                ���茋��
' =                                                 1) �t�@�C��
' =                                                 2) �t�H���_�[
' =                                                 0) �G���[�i���݂��Ȃ��p�X�j
' = �o��    FileSystemObject ���g���Ă���̂ŁA�t�@�C��/�t�H���_��
' =         ���݊m�F�ɂ��g�p�\�B
' = �ˑ�    �Ȃ�
' = ����    FileSystem.vbs
' ==================================================================
Public Function GetFileOrFolder( _
    ByVal sChkTrgtPath _
)
    Dim oFileSys
    Dim bFolderExists
    Dim bFileExists
    
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    bFolderExists = oFileSys.FolderExists(sChkTrgtPath)
    bFileExists = oFileSys.FileExists(sChkTrgtPath)
    Set oFileSys = Nothing
    
    If bFolderExists = False And bFileExists = True Then
        GetFileOrFolder = 1 '�t�@�C��
    ElseIf bFolderExists = True And bFileExists = False Then
        GetFileOrFolder = 2 '�t�H���_�[
    Else
        GetFileOrFolder = 0 '�G���[�i���݂��Ȃ��p�X�j
    End If
End Function

