Option Explicit

'�t�@�C�����w�肷��ƌ��ݎ�����t�^�����o�b�N�A�b�v�t�@�C�����쐬����B
'�����t�@�C�����̂��̂����݂��Ă�����A�A���t�@�x�b�g��t�^�����o�b�N�A�b�v�t�@�C�����쐬����B
'	ex. 211201a, 211202b, �c
'�w�萔���o�b�N�A�b�v�����܂�����A�Â����̂���폜����B
'�o�b�N�A�b�v�Ώۂ̓t�@�C���̂݁B
'�������Ƀo�b�N�A�b�v�t�@�C�������w��ł���B
'���v�C�����o�b�N�A�b�v�ő吔���A�������t�̃o�b�N�A�b�v�t�@�C���Ŗ��������ƁA�V�����t�@�C�����X�V����Ă������ߗv���ӁB

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )		'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()

'===============================================================================
'= �ݒ�l
'===============================================================================
Const lMAX_BAK_FILE_NUM_DEFAULT = 50
Const sBAK_DIR_NAME = "_bak"
Const sBAK_FILE_SUFFIX = "#b#"

'===============================================================================
'= �{����
'===============================================================================
Const sSCRIPT_NAME = "�t�@�C���o�b�N�A�b�v"
Dim sTrgtFilePath
Dim lBakFileNumMax
If WScript.Arguments.Count >= 2 Then
	sTrgtFilePath = WScript.Arguments(0)
	lBakFileNumMax = CLng(WScript.Arguments(1))
ElseIf WScript.Arguments.Count = 1 Then
	sTrgtFilePath = WScript.Arguments(0)
	lBakFileNumMax = lMAX_BAK_FILE_NUM_DEFAULT
Else
	WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
	WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'�Ώۃt�@�C�����擾
Dim sTrgtFileParDirPath
Dim sTrgtFileBaseName
Dim sTrgtFileExt
Dim sDateSuffix
sTrgtFileParDirPath = objFSO.GetParentFolderName( sTrgtFilePath )
sTrgtFileBaseName = objFSO.GetBaseName( sTrgtFilePath )
sTrgtFileExt = objFSO.GetExtensionName( sTrgtFilePath )
sDateSuffix = ConvDate2String(Now(),2)

'�o�b�N�A�b�v�t�@�C�����쐬
Dim sBakDirPath
Dim sBakFilePathBase
Dim sBakFilePath
sBakDirPath = sTrgtFileParDirPath & "\" & sBAK_DIR_NAME
sBakFilePathBase = sBakDirPath & "\" & sTrgtFileBaseName & "_" & sBAK_FILE_SUFFIX
sBakFilePath = sBakFilePathBase & sDateSuffix & "." & sTrgtFileExt

'�o�b�N�A�b�v�t�H���_�쐬
Call CreateDirectry( sBakDirPath )

'*** �t�@�C���o�b�N�A�b�v ***
'�����݃t�@�C���p�X����
Dim lAlphaIdx
lAlphaIdx = 97 'ascii�R�[�h��a
Do While objFSO.FileExists( sBakFilePath )
	sBakFilePath = sBakFilePathBase & sDateSuffix & Chr(lAlphaIdx) & "." & sTrgtFileExt
	lAlphaIdx = lAlphaIdx + 1
Loop
'�t�@�C���o�b�N�A�b�v
objFSO.CopyFile sTrgtFilePath, sBakFilePath

'*** �Â��t�@�C���폜 ***
'�t�@�C�����X�g�擾
Dim cFileList
Set cFileList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct( sBakDirPath, cFileList, 1, "*")

'�o�b�N�A�b�v�t�@�C�����擾
Dim lBakFileNum
Dim sDelFilePath
lBakFileNum = 0
sDelFilePath = ""
Dim sFilePath
For Each sFilePath in cFileList
	If ( (InStr(sFilePath, sBakFilePathBase) > 0) And _
	     (objFSO.GetExtensionName(sFilePath) = sTrgtFileExt) ) Then
		 If lBakFileNum = 0 Then
			sDelFilePath = sFilePath
		 End If
		 lBakFileNum = lBakFileNum + 1
	End If
Next

'�o�b�N�A�b�v�t�@�C���폜
If lBakFileNum > lBakFileNumMax Then
	objFSO.DeleteFile sDelFilePath, True
End If

'MsgBox "�o�b�N�A�b�v�����I", vbOKOnly, sSCRIPT_NAME

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
