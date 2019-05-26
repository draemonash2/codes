Option Explicit

'===============================================================================
'�y�����z
'	�������OCSV�𐮌`���ACANape�ŃC���|�[�g�ł���`���ɕϊ�����B
'	���`������e�͈ȉ��̒ʂ�B
'		- �uDatatype�v���t�^
'			�uDataType�v�� data_type_list.csv ���擾����B
'		- RAM����u������iex. ram[0]:1 �� ram_0:1�j
'	data_type_list.csv �����݂��Ȃ��ꍇ�́A���ׂ� uint8 �Ɖ��߂��ĕϊ�����B
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
'	1.0.0	2019/05/13	�E�V�K�쐬
'	1.1.0	2019/05/26	�E�������OCSV�o�b�N�A�b�v�o�͏����ύX
'						�E���̑����t�@�N�^�����O
'===============================================================================

'===============================================================================
'= �C���N���[�h
'===============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\..\_lib\String.vbs" )		'GetFileExt()
Call Include( sMyDirPath & "\..\_lib\FileSystem.vbs" )	'GetFileList3()
Call Include( sMyDirPath & "\..\_lib\Collection.vbs" )	'ReadTxtFileToCollection()
														'WriteTxtFileFrCollection()
Call Include( sMyDirPath & "\..\_lib\String.vbs" )		'GetFileNotExistPath()

'===============================================================================
' �ݒ�
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"
CONST CREATE_BACKUP_FILE = False
CONST DEFAULT_DATA_TYPE = "uint8"

'===============================================================================
' �{����
'===============================================================================
Const RAMNAME_ROW_KEYWORD = "TimeStamp"
Const DATATYPE_ROW_KEYWORD = "DataType"

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
	Set cFileList = Nothing
ElseIf WScript.Arguments.Count = 1 And _
	objFSO.FileExists(WScript.Arguments(0)) Then
	cCsvFileList.add WScript.Arguments(0)
Else
	WScript.Echo "�����G���["
	WScript.Quit
End If

'*****************************
' DataType�ꗗ�擾
'*****************************
dim dDataTypeList
Set dDataTypeList = CreateObject("Scripting.Dictionary")

Dim sDataTypeListFilePath
sDataTypeListFilePath = sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME

dim objTxtFile
If objFSO.FileExists(sDataTypeListFilePath) Then
	set objTxtFile = objFSO.OpenTextFile(sDataTypeListFilePath, 1)

	dim objWords
	Dim sTxtLine
	Do Until objTxtFile.AtEndOfStream
		sTxtLine = objTxtFile.ReadLine
		objWords = split(sTxtLine, ",")
		objWords(0) = ReplaceKeyword(objWords(0))
		On Error Resume Next '�d���L�[���������疳��
		dDataTypeList.Add objWords(0), objWords(1) 'RamName DataType
		On Error Goto 0
	Loop
	objTxtFile.Close
Else
	'Do Nothing
End If

'*****************************
' �������OCSV���`
'*****************************
dim sCsvFilePath
for each sCsvFilePath In cCsvFileList
	
	'*** �������OCSV�I�[�v�� ***
	dim cFileContents
	Set cFileContents = CreateObject("System.Collections.ArrayList")
	call ReadTxtFileToCollection(sCsvFilePath, cFileContents)
	
	'*** �������O�t�@�C���`�F�b�N ***
	If Left(cFileContents(0), len(RAMNAME_ROW_KEYWORD)) = RAMNAME_ROW_KEYWORD Then
		
		'*** �o�b�N�A�b�v�o�� ***
		If CREATE_BACKUP_FILE = True then
			Dim sCsvBakFilePath
			sCsvBakFilePath = sCsvFilePath & ".bak"
			sCsvBakFilePath = GetFileNotExistPath(sCsvBakFilePath)
			objFSO.CopyFile sCsvFilePath, sCsvBakFilePath
		End If
		
		'*** �ϐ����u�� ***
		cFileContents(0) = ReplaceKeyword(cFileContents(0))
		
		'*** Datatype�u��or�}�� ***
		Dim vRamNames
		vRamNames = Split(cFileContents(0), ",")
		Dim sRamName
		Dim sDataTypeLine
		Dim lIdx
		lIdx = 0
		for each sRamName In vRamNames
			If lIdx = 0 Then '1��ڂ͖���
				sDataTypeLine = DATATYPE_ROW_KEYWORD
			else
				'sRamName = ReplaceKeyword(sRamName) '���łɒu���ς�
				if dDataTypeList.Exists(sRamName) then
					sDataTypeLine = sDataTypeLine & "," & dDataTypeList.Item(sRamName)
				else
					sDataTypeLine = sDataTypeLine & "," & DEFAULT_DATA_TYPE
				end if
			end if
			lIdx = lIdx + 1
		next
		If Left(cFileContents(1), len(DATATYPE_ROW_KEYWORD)) = DATATYPE_ROW_KEYWORD Then
			cFileContents(1) = sDataTypeLine
		Else
			cFileContents.Insert 1, sDataTypeLine
		End If
		
		'*** CSV�o�� ***
		call WriteTxtFileFrCollection(sCsvFilePath, cFileContents, True)
	Else
		'Do Nothing
	End If
	
	Set cFileContents = Nothing
next

Set objFSO = Nothing
Set cCsvFileList = Nothing
Set dDataTypeList = Nothing

MsgBox "�������OCSV ���`����!"

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
