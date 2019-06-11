Option Explicit

'===============================================================================
'�y�����z
'	�������OCSV�𐮌`���ACANape�ŃC���|�[�g�ł���`���ɕϊ�����B
'	���`������e�͈ȉ��̒ʂ�B
'		- �uDatatype�v���t�^
'			�uDataType�v�� data_type_list.csv ���擾����B
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
'	0.1.0	2019/05/13	�E�V�K�쐬
'	0.1.1	2019/05/26	�E�������OCSV�o�b�N�A�b�v�o�͏����ύX
'						�E���̑����t�@�N�^�����O
'	0.2.0	2019/06/02	�E�v���O���X�o�[����
'	0.3.0	2019/06/11	�E�z��w��[]�L���u�������폜
'===============================================================================

'===============================================================================
'= �C���N���[�h
'===============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\String.vbs" )				'GetFileExt()
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )			'GetFileList3()
Call Include( "C:\codes\vbs\_lib\Collection.vbs" )			'ReadTxtFileToCollection()
															'WriteTxtFileFrCollection()
Call Include( "C:\codes\vbs\_lib\String.vbs" )				'GetFileNotExistPath()
Call Include( "C:\codes\vbs\_lib\ProgressBarCscript.vbs" )	'Class ProgressBar

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

Dim objPrgrsBar
Set objPrgrsBar = New ProgressBar
objPrgrsBar.Message = "�������OCSV���`��..."

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
Dim lProcIdx
Dim lProcNum
lProcIdx = 0
lProcNum = cCsvFileList.Count
Call objPrgrsBar.Update(lProcIdx, lProcNum)
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
	
	lProcIdx = lProcIdx + 1
	Call objPrgrsBar.Update(lProcIdx, lProcNum)
	
	Set cFileContents = Nothing
next

Set objFSO = Nothing
Set cCsvFileList = Nothing
Set dDataTypeList = Nothing

MsgBox "�������OCSV ���`����!"

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
