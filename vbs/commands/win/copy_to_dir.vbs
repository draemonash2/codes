Option Explicit

'==============================================================================
'�y�����z
'	�t�@�C��/�t�H���_���R�s�[����B
'	�ړ���̃t�H���_�����݂��Ȃ��ꍇ�A�t�H���_���쐬���Ă���R�s�[����B
'
'�y�g�p���@�z
'	copy_to_dir.vbs <source_path> <destination_path>
'
'�y�g�p��z
'	1) copy_to_dir.vbs c:\codes\vbs\test.txt c:\test\test.txt
'	2) copy_to_dir.vbs c:\codes\vbs c:\test\vbs
'		c:\codes\vbs
'			�� a.txt
'			�� b
'				�� c.txt
'		��
'		c:\test\vbs
'			�� a.txt
'			�� b
'				�� c.txt
'
'�y�o�������z
'	�Ȃ�
'
'�y���������z
'	1.0.0	2019/05/12	�V�K�쐬
'==============================================================================

'==============================================================================
' �ݒ�
'==============================================================================

'==============================================================================
'= �C���N���[�h
'==============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\String.vbs" )			'GetDirPath()
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )		'CreateDirectry()
														'GetFileOrFolder()

'==============================================================================
' �{����
'==============================================================================
'�����`�F�b�N
If WScript.Arguments.Count = 2 Then
	'Do Nothing
Else
	Wscript.quit
End If

dim sSrcPath
dim sDstPath
sSrcPath = Replace(WScript.Arguments(0), "/", "\")
sDstPath = Replace(WScript.Arguments(1), "/", "\")

Dim lSrcPathType
lSrcPathType = GetFileOrFolder(sSrcPath)

dim sDstParDir
sDstParDir = GetDirPath( sDstPath )

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If lSrcPathType = 1 Then '�t�@�C��
	call CreateDirectry( sDstParDir )
	objFSO.CopyFile sSrcPath, sDstPath
ElseIf lSrcPathType = 2 Then '�t�H���_
	call CreateDirectry( sDstParDir )
	objFSO.CopyFolder sSrcPath, sDstPath
Else '������
'	WScript.Echo "�t�@�C�������݂��܂���"
End If

Set objFSO = Nothing

'==============================================================================
'= �C���N���[�h�֐�
'==============================================================================
' �O���v���O���� �C���N���[�h�֐�
Private Function Include( _
	ByVal sOpenFile _
)
	Dim objFSO
	Dim objVbsFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	sOpenFile = objFSO.GetAbsolutePathName( sOpenFile )
	Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
	
	ExecuteGlobal objVbsFile.ReadAll()
	objVbsFile.Close
	
	Set objVbsFile = Nothing
	Set objFSO = Nothing
End Function
