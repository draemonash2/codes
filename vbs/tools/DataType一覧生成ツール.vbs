Option Explicit

'===============================================================================
'�y�����z
'	�������OCSV����DataType�ꗗ�𐶐�����
'
'�y�g�p���@�z
'	�g�p���@�͂Q�ʂ�B
'		���t�H���_�z���̑Scsv���ׂĂ�u���������ꍇ
'			1. �{�X�N���v�g��u���������t�H���_�Ɉړ�����B
'			2. �{�X�N���v�g�����s����B(�_�u���N���b�N)
'		���P�t�@�C���̂ݐ��`�������ꍇ
'			1. ���`�������������OCSV��{�X�N���v�g��drag&drop����
'
'�y�o�������z
'	�Ȃ�
'
'�y���������z
'	0.1.0	2019/05/12	�x�[�^�ŁiTODO:�v�d���폜�����j
'===============================================================================

'===============================================================================
'= �C���N���[�h
'===============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\..\_lib\FileSystem.vbs" )	'GetFileList3()
Call Include( sMyDirPath & "\..\_lib\Collection.vbs" )	'ReadTxtFileToCollection()
														'WriteTxtFileFrCollection()

'===============================================================================
' �ݒ�
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"

'===============================================================================
' �{����
'===============================================================================
Const DATA_ROW_KEYWORD = "DataType"
Const RAMNAME_ROW_KEYWORD = "Data"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'*****************************
' �������OCSV�t�@�C�����X�g�擾
'*****************************
dim cCsvFileList
Set cCsvFileList = CreateObject("System.Collections.ArrayList")
Dim sRootDirPath
sRootDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )
If WScript.Arguments.Count = 0 Then
	dim cFileList
	Set cFileList = CreateObject("System.Collections.ArrayList")
	call GetFileList3(sRootDirPath, cFileList, 1)

	dim sFilePath
	for each sFilePath in cFileList
		if objFSO.GetExtensionName(sFilePath) = "csv" And _
		   objFSO.GetFileName(sFilePath) <> DATA_TYPE_LIST_FILE_NAME then
			cCsvFileList.add sFilePath
		end if
	next
ElseIf WScript.Arguments.Count = 1 And _
	objFSO.FileExists(WScript.Arguments(0)) Then
	cCsvFileList.add WScript.Arguments(0)
Else
	WScript.Echo "�����G���["
	WScript.Quit
End If

'*****************************
' DataType�擾
'*****************************
dim oDataTypeList
Set oDataTypeList = CreateObject("System.Collections.ArrayList")
dim sCsvFilePath
for each sCsvFilePath In cCsvFileList
	'*** �������OCSV�I�[�v�� ***
	dim cFileContents
	Set cFileContents = CreateObject("System.Collections.ArrayList")
	call ReadTxtFileToCollection(sCsvFilePath, cFileContents)

	'*** DataType�擾 ***
	If InStr(cFileContents(1), DATA_ROW_KEYWORD) Then
		Dim vRamNames
		Dim vDataTypes
		vRamNames = Split(cFileContents(0), ",")
		vDataTypes = Split(cFileContents(1), ",")
		Dim sRamNameRaw
		Dim lIdx
		lIdx = 0
		for each sRamNameRaw In vRamNames
			If sRamNameRaw = RAMNAME_ROW_KEYWORD Then
				'Do Nothing
			else
				oDataTypeList.Add sRamNameRaw & "," & vDataTypes(lIdx)
			end if
			lIdx = lIdx + 1
		next
	Else
		'Do Nothing
	End If
next

'*****************************
' DataType�ꗗ�o��
'*****************************
call WriteTxtFileFrCollection(sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME, oDataTypeList)

MsgBox "DataType�ꗗ�o�͊���!"

'===============================================================================
' �֐�
'===============================================================================
Private Function ReplaceKeyword( _
	byval sTrgtWord _
)
	Dim sOutWord
	sOutWord = sTrgtWord
	sOutWord = Replace(sOutWord, "[", "_")
	sOutWord = Replace(sOutWord, "]", "")
	ReplaceKeyword = sOutWord
End Function

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
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
