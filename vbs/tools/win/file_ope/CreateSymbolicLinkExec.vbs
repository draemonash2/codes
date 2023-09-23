'Option Explicit

'�y���ӎ����z
'   �E�{�X�N���v�g�����s����ɂ́A���K�w�ɁuCreateSymbolicLink.vbs�v���i�[����K�v������B
'     
'     <�{�X�N���v�g CreateSymbolicLinkExec.vbs ��݂������R>
'       �V���{���b�N�����N�̍쐬�͊Ǘ��Ҍ������K�{�B
'       X-Finder����Ǘ��Ҍ����Ŏ��s����ɂ́A�V���{���b�N�����N���쐬���鏈����
'       �ʂ̃X�N���v�g�iCreateSymbolicLinkExec.vbs�j�Ƃ��Đ؂�o���A���̃X�N���v�g��
'       �Ǘ��Ҍ����Ƃ��Ď��s����K�v������B
'       �Ȃ��A�Ǘ��Ҍ����Ŏ��s����ꍇ�͈�����n���Ȃ����߁A�������e�L�X�g�t�@�C���Ƃ���
'       �����o���Ă���A�Ăяo���B

'####################################################################
'### �ݒ�
'####################################################################
Const OBJECT_SUFFIX = ".symlink"
Const sARG_FILE_NAME = "CreateSymbolicLinkExecArg.txt" '���O�́uCreateSymbolicLink.vbs�v���̐ݒ�l�ƍ��킹�邱��

'####################################################################
'### ���O����
'####################################################################
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" )    'ExecDosCmd()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileOrFolder()

'####################################################################
'### �{����
'####################################################################
Const sPROG_NAME = "�V���{���b�N�����N�쐬�i�������j"

'*** �Ώۃt�@�C��/�t�H���_���擾 ***
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sTrgtFilePath
sTrgtFilePath = objFSO.GetSpecialFolder(2) & "\" & sARG_FILE_NAME
Dim objTxtFile
Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
DIm cFilePaths
Set cFilePaths = CreateObject("System.Collections.ArrayList")
Do Until objTxtFile.AtEndOfStream
    cFilePaths.Add objTxtFile.ReadLine
Loop
objTxtFile.Close
objFSO.DeleteFile sTrgtFilePath, True
'������debug������
'Dim sArg
'For Each sArg In cFilePaths
'    MsgBox sArg, vbYes, sPROG_NAME
'Next
'WScript.Quit
'������debug������

'*** �t�@�C���p�X�`�F�b�N ***
If cFilePaths.Count = 0 Then
    MsgBox "�I�u�W�F�N�g���I������Ă��܂���", vbYes, sPROG_NAME
    MsgBox "�����𒆒f���܂�", vbYes, sPROG_NAME
    WScript.Quit
End If

'*** �V���{���b�N�����N�쐬 ***
Dim oObjPath
Dim lObjType '0:notexists 1:file 2:folder
For Each oObjPath In cFilePaths
    lObjType = GetFileOrFolder( oObjPath )
    
    Dim sDstPath
    Dim sSrcPath
    sDstPath = oObjPath
    Dim sCmd
    If lObjType = 1 Then 'file
       'sSrcPath = objFSO.GetParentFolderName( oObjPath ) & "\" & _
       '           objFSO.GetBaseName( oObjPath ) & OBJECT_SUFFIX & "." & _
       '           objFSO.GetExtensionName( oObjPath )
        sSrcPath = oObjPath & OBJECT_SUFFIX & "." & objFSO.GetExtensionName( oObjPath )
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

'MsgBox "�V���{���b�N�����N���쐬���܂���", vbYes, sPROG_NAME

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

