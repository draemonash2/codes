'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'�y���ӎ����z
'   �E�{�X�N���v�g�����s����ɂ́A���K�w�ɁuCreateSymbolicLinkExec.vbs�v���i�[����K�v������B
'     
'     <CreateSymbolicLinkExec.vbs ��݂������R>
'       �V���{���b�N�����N�̍쐬�͊Ǘ��Ҍ������K�{�B
'       X-Finder����Ǘ��Ҍ����Ŏ��s����ɂ́A�V���{���b�N�����N���쐬���鏈����
'       �ʂ̃X�N���v�g�iCreateSymbolicLinkExec.vbs�j�Ƃ��Đ؂�o���A���̃X�N���v�g��
'       �Ǘ��Ҍ����Ƃ��Ď��s����K�v������B
'       �Ȃ��A�Ǘ��Ҍ����Ŏ��s����ꍇ�͈�����n���Ȃ����߁A�������e�L�X�g�t�@�C���Ƃ���
'       �����o���Ă���A�Ăяo���B

'####################################################################
'### �ݒ�
'####################################################################
Const sARG_FILE_NAME = "CreateSymbolicLinkExecArg.txt" '���O�́uCreateSymbolicLinkExec.vbs�v���̐ݒ�l�ƍ��킹�邱��
Const sEXEC_SCRIPT_FILE_NAME = "CreateSymbolicLinkExec.vbs"

'####################################################################
'### ���O����
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" )    'ExecRunas2()

'####################################################################
'### �{����
'####################################################################
Const sPROG_NAME = "�V���{���b�N�����N�쐬�i���O�����j"

'*** �Ώۃt�@�C��/�t�H���_�p�X�擾 ***
DIm cFilePaths
Dim sCurDirPath
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
    sCurDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
    Set cFilePaths = WScript.Col( WScript.Env("Selected") )
    sCurDirPath = WScript.Env( "%MYDIRPATH_CODES%\vbs\tools\win\file_ope" )
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
    sCurDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
End If
'������debug������
'For Each sArg In cFilePaths
'    MsgBox sArg, vbYes, sPROG_NAME
'Next
'������debug������

'*** �Ώۃt�@�C��/�t�H���_�p�X���o�� ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sTrgtFilePath
sTrgtFilePath = objFSO.GetSpecialFolder(2) & "\" & sARG_FILE_NAME
Dim objTxtFile
Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
Dim oObjPath
For Each oObjPath In cFilePaths
    'MsgBox oObjPath, vbYes, sPROG_NAME
    objTxtFile.WriteLine oObjPath
Next
objTxtFile.Close

'*** �Ǘ��Ҍ����ŃX�N���v�g���s ***
Dim sScriptPath
sScriptPath = sCurDirPath & "\" & sEXEC_SCRIPT_FILE_NAME
Call ExecRunas2(sScriptPath)

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

