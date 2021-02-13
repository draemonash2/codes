'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'�y���ӎ����zX-Finder������s����ꍇ�́A�Ǘ��Ҍ����ɂċN������X-Finder����N�����邱��

'####################################################################
'### �ݒ�
'####################################################################
Const OBJECT_SUFFIX = " - �V���{���b�N�����N"

'####################################################################
'### ���O����
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" )    'ExecRunas()
                                                            'ExecDosCmd()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileOrFolder()

'####################################################################
'### �{����
'####################################################################
Const sPROG_NAME = "�V���{���b�N�����N�쐬"

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
    Call ExecRunas() '�Ǘ��҂Ƃ��Ď��s
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
    Call ExecRunas() '�Ǘ��҂Ƃ��Ď��s
End If
'������debug������
'For Each sArg In cFilePaths
'    msgbox sArg
'Next
'������debug������

'*** �t�@�C���p�X�`�F�b�N ***
If cFilePaths.Count = 0 Then
    MsgBox "�I�u�W�F�N�g���I������Ă��܂���", vbYes, sPROG_NAME
    MsgBox "�����𒆒f���܂�", vbYes, sPROG_NAME
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
        MsgBox "�I�u�W�F�N�g�����݂��܂���", vbYes, sPROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, sPROG_NAME
        WScript.Quit
    End If
    '������debug������
    'msgbox sCmd
    '������debug������
    call ExecDosCmd( sCmd )
Next

MsgBox "�V���{���b�N�����N���쐬���܂���", vbYes, sPROG_NAME

'####################################################################
'### �C���N���[�h�֐�
'####################################################################
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

