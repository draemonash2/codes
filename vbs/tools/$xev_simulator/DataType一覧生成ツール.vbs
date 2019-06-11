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
'	0.0.0	2019/05/13	�E�V�K�쐬
'	0.1.0	2019/05/26	�EDataType�ꗗ�̃o�b�N�A�b�v�쐬��I���ł���悤�ɕύX
'						�E�������OCSV�t�@�C�����菈���ύX
'	0.2.0	2019/06/11	�E�z��w��[]�L���u�������폜
'						�E�v���O���X�o�[����
'===============================================================================

'===============================================================================
'= �C���N���[�h
'===============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
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

'===============================================================================
' �{����
'===============================================================================
Const RAMNAME_ROW_KEYWORD = "TimeStamp"
Const DATATYPE_ROW_KEYWORD = "DataType"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim objPrgrsBar
Set objPrgrsBar = New ProgressBar
objPrgrsBar.Message = "�������OCSV���`��..."

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
' DataType�擾
'*****************************
dim oDataTypeList
Set oDataTypeList = CreateObject("System.Collections.ArrayList")
dim dDataTypeListDupChk '�d���`�F�b�N�p
set dDataTypeListDupChk = CreateObject("Scripting.Dictionary")
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
	If Left(cFileContents(0), len(RAMNAME_ROW_KEYWORD)) = RAMNAME_ROW_KEYWORD And _
	   Left(cFileContents(1), len(DATATYPE_ROW_KEYWORD)) = DATATYPE_ROW_KEYWORD Then
		
		'*** �o�b�N�A�b�v�o�� ***
		If CREATE_BACKUP_FILE = True then
			Dim sCsvBakFilePath
			sCsvBakFilePath = sCsvFilePath & ".bak"
			sCsvBakFilePath = GetFileNotExistPath(sCsvBakFilePath)
			objFSO.CopyFile sCsvFilePath, sCsvBakFilePath
		End If
		
		'*** DataType�擾 ***
		Dim vRamNames
		Dim vDataTypes
		vRamNames = Split(cFileContents(0), ",")
		vDataTypes = Split(cFileContents(1), ",")
		Dim sRamName
		Dim lIdx
		lIdx = 0
		for each sRamName In vRamNames
			If lIdx = 0 Then '1��ڂ͖���
				'Do Nothing
			else
				Dim sDataTypeListLine
				sDataTypeListLine = sRamName & "," & vDataTypes(lIdx)
				If Not dDataTypeListDupChk.Exists( sDataTypeListLine ) Then
					oDataTypeList.Add sDataTypeListLine
					dDataTypeListDupChk.Add sDataTypeListLine, ""
				end if
			end if
			lIdx = lIdx + 1
		next
	Else
		'Do Nothing
	End If
	
	lProcIdx = lProcIdx + 1
	Call objPrgrsBar.Update(lProcIdx, lProcNum)
	
	Set cFileContents = Nothing
next

'*****************************
' DataType�ꗗ�o��
'*****************************
call WriteTxtFileFrCollection(sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME, oDataTypeList, True)

Set objFSO = Nothing
Set cCsvFileList = Nothing
Set oDataTypeList = Nothing
set dDataTypeListDupChk = Nothing

MsgBox "DataType�ꗗ ��������!"

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
