Option Explicit

'==============================================================================
' �ݒ�
'==============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"

'==============================================================================
' �{����
'==============================================================================
Const DATA_ROW_KEYWORD = "DataType"
Const RAMNAME_ROW_KEYWORD = "Data"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'*****************************
' DataType�ꗗ�擾
'*****************************
dim oDataTypeList
set oDataTypeList = CreateObject("Scripting.Dictionary")

Dim sRootDirPath
sRootDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )

Dim sDataTypeListFilePath
sDataTypeListFilePath = sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME

dim objTxtFile
set objTxtFile = objFSO.OpenTextFile(sDataTypeListFilePath, 1, True)

dim objWords
Dim sTxtLine
Do Until objTxtFile.AtEndOfStream
	sTxtLine = objTxtFile.ReadLine
	objWords = split(sTxtLine, ",")
	if InStr(objWords(0), "[") Then
		objWords(0) = ReplaceKeyword(objWords(0))
	Else
		'Do Nothing
	end if
	oDataTypeList.Add objWords(0), objWords(1)
Loop
objTxtFile.Close

'*****************************
' ����csv�t�@�C�����X�g�擾
'*****************************
dim cFileList
Set cFileList = CreateObject("System.Collections.ArrayList")
call GetFileList2(sRootDirPath, cFileList, 1)

dim cCsvFileList
Set cCsvFileList = CreateObject("System.Collections.ArrayList")
dim sFilePath
for each sFilePath in cFileList
	if objFSO.GetExtensionName(sFilePath) = "csv" And _
	   objFSO.GetFileName(sFilePath) <> DATA_TYPE_LIST_FILE_NAME then
		cCsvFileList.add sFilePath
	end if
next

'*****************************
' ��������csv���`
'*****************************
dim sCsvFilePath
for each sCsvFilePath In cCsvFileList
	'*** ����csv�o�b�N�A�b�v�o�� ***
	Dim sCsvBakFilePath
	sCsvBakFilePath = sCsvFilePath & ".bak"
	If objFSO.FileExists(sCsvBakFilePath) Then
		'Do Nothing
	Else
		objFSO.CopyFile sCsvFilePath, sCsvBakFilePath
	End If

	'*** ����csv�I�[�v�� ***
	dim cFileContents
	Set cFileContents = CreateObject("System.Collections.ArrayList")
	call ReadTxtFileToArray(sCsvFilePath, cFileContents)

	'*** �ϐ����u�� ***
	cFileContents(0) = ReplaceKeyword(cFileContents(0))

	'*** Datatype�}�� ***
	If InStr(cFileContents(1), DATA_ROW_KEYWORD) Then
		'Do Nothing
	Else
		Dim vRamNames
		vRamNames = Split(cFileContents(0), ",")
		Dim sRamNameRaw
		Dim sRamNameRep
		Dim sDataTypeLine
		sDataTypeLine = DATA_ROW_KEYWORD
		for each sRamNameRaw In vRamNames
			If sRamNameRaw = RAMNAME_ROW_KEYWORD Then
				'Do Nothing
			else
				sRamNameRep = sRamNameRaw
				sRamNameRep = ReplaceKeyword(sRamNameRep)
				if oDataTypeList.Exists(sRamNameRaw) then
					sDataTypeLine = sDataTypeLine & "," & oDataTypeList.Item(sRamNameRaw)
				elseif oDataTypeList.Exists(sRamNameRep) then
					sDataTypeLine = sDataTypeLine & "," & oDataTypeList.Item(sRamNameRep)
				else
					sDataTypeLine = sDataTypeLine & ",uint8"
				end if
			end if
		next
		cFileContents.Insert 1, sDataTypeLine
	End If

	'*** csv�o�� ***
	call WriteTxtFileFrArray(sCsvFilePath, cFileContents)
next

MsgBox "��������CSV���`����!"

'==============================================================================
' �֐�
'==============================================================================
Private Function ReplaceKeyword( _
	byval sTrgtWord _
)
	Dim sOutWord
	sOutWord = sTrgtWord
	sOutWord = Replace(sOutWord, "[", "_")
	sOutWord = Replace(sOutWord, "]", "")
	ReplaceKeyword = sOutWord
End Function

'==============================================================================
' ���C�u����
'==============================================================================
' ==================================================================
' = �T�v	�t�@�C��/�t�H���_�p�X�ꗗ���擾����
' = ����	sTrgtDir		String		[in]	�Ώۃt�H���_
' = ����	cFileList		Collections [out]	�t�@�C��/�t�H���_�p�X�ꗗ
' = ����	lFileListType	Long		[in]	�擾����ꗗ�̌`��
' =													0�F����
' =													1:�t�@�C��
' =													2:�t�H���_
' =													����ȊO�F�i�[���Ȃ�
' = �ߒl	�Ȃ�
' = �o��	�EDir �R�}���h�ɂ��t�@�C���ꗗ�擾�BGetFileList() ���������B
' ==================================================================
Public Function GetFileList2( _
	ByVal sTrgtDir, _
	ByRef cFileList, _
	ByVal lFileListType _
)
	Dim objFSO	'FileSystemObject�̊i�[��
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	
	'Dir �R�}���h���s�i�o�͌��ʂ��ꎞ�t�@�C���Ɋi�[�j
	Dim sTmpFilePath
	Dim sExecCmd
	sTmpFilePath = WScript.CreateObject( "WScript.Shell" ).CurrentDirectory & "\Dir.tmp"
	Select Case lFileListType
		Case 0:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a > """ & sTmpFilePath & """"
		Case 1:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:a-d > """ & sTmpFilePath & """"
		Case 2:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:d > """ & sTmpFilePath & """"
		Case Else: sExecCmd = ""
	End Select
	With CreateObject("Wscript.Shell")
		.Run "cmd /c" & sExecCmd, 7, True
	End With
	
	Dim objFile
	Dim sTextAll
	On Error Resume Next
	If Err.Number = 0 Then
		Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
		If Err.Number = 0 Then
			sTextAll = objFile.ReadAll
			sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '�����ɉ��s���t�^����Ă��܂����߁A�폜
			Dim vFileList
			vFileList = Split( sTextAll, vbNewLine )
			Dim sFilePath
			For Each sFilePath In vFileList
				cFileList.add sFilePath
			Next
			objFile.Close
		Else
			WScript.Echo "�t�@�C�����J���܂���: " & Err.Description
		End If
		Set objFile = Nothing	'�I�u�W�F�N�g�̔j��
	Else
		WScript.Echo "�G���[ " & Err.Description
	End If	
	objFSO.DeleteFile sTmpFilePath, True
	Set objFSO = Nothing	'�I�u�W�F�N�g�̔j��
	On Error Goto 0
End Function
'	Call Test_GetFileList2()
	Private Sub Test_GetFileList2()
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Dim sCurDir
		sCurDir = objFSO.GetParentFolderName( WScript.ScriptFullName )
		
		msgbox sCurDir
		
		Dim cFileList
		Set cFileList = CreateObject("System.Collections.ArrayList")
		Call GetFileList2( sCurDir, cFileList, 1 )
		
		dim sFilePath
		dim sOutput
		sOutput = ""
		for each sFilePath in cFileList
			sOutput = sOutput & vbNewLine & sFilePath
		next
		MsgBox sOutput
	End Sub

' ==================================================================
' = �T�v	�w�肳�ꂽ�t�@�C���p�X����g���q�𒊏o����B
' = ����	sFilePath	String	[in]  �t�@�C���p�X
' = �ߒl				String		  �g���q
' = �o��	�E�g���q���Ȃ��ꍇ�A�󕶎���ԋp����
' =			�E�t�@�C�������w��\
' ==================================================================
Public Function GetFileExt( _
	ByVal sFilePath _
)
	Dim sFileName
	sFileName = GetFileName(sFilePath)
	If InStr(sFileName, ".") > 0 Then
		GetFileExt = ExtractTailWord(sFileName, ".")
	Else
		GetFileExt = ""
	End If
End Function
	'Call Test_GetFileExt()
	Private Sub Test_GetFileExt()
		Dim Result
		Result = "[Result]"
		Result = Result & vbNewLine & GetFileExt("c:\codes\test.txt")	  'txt
		Result = Result & vbNewLine & GetFileExt("c:\codes\test")		  '
		Result = Result & vbNewLine & GetFileExt("test.txt")			  'txt
		Result = Result & vbNewLine & GetFileExt("test")				  '
		Result = Result & vbNewLine & GetFileExt("c:\codes\test.aaa.txt") 'txt
		Result = Result & vbNewLine & GetFileExt("test.aaa.txt")		  'txt
		MsgBox Result
	End Sub

' ==================================================================
' = �T�v	�e�L�X�g�t�@�C���̒��g��z��Ɋi�[
' = ����	sTrgtFilePath	String		[in]	�t�@�C���p�X
' = ����	cFileContents	Collections [out]	�t�@�C���̒��g
' = �ߒl	�Ȃ�
' = �o��	�Ȃ�
' ==================================================================
Public Function ReadTxtFileToArray( _
	ByVal sTrgtFilePath, _
	ByRef cFileContents _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTxtFile
	Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
	
	Do Until objTxtFile.AtEndOfStream
		cFileContents.add objTxtFile.ReadLine
	Loop
	
	objTxtFile.Close
End Function
'	Call Test_OpenTxtFile2Array()
	Private Sub Test_OpenTxtFile2Array()
		Dim cFileList
		Set cFileList = CreateObject("System.Collections.ArrayList")
		sFilePath = "C:\codes\vbs\��������CSV���`�c�[��\data_type_list.csv"
		call ReadTxtFileToArray( sFilePath, cFileList )
		
		dim sFilePath
		dim sOutput
		sOutput = ""
		for each sFilePath in cFileList
			sOutput = sOutput & vbNewLine & sFilePath
		next
		MsgBox sOutput
	End Sub

' ==================================================================
' = �T�v	�z��̒��g���e�L�X�g�t�@�C���ɏ����o��
' = ����	sTrgtFilePath	String		[in]	�t�@�C���p�X
' = ����	cFileContents	Collections [in]	�t�@�C���̒��g
' = �ߒl	�Ȃ�
' = �o��	�Ȃ�
' ==================================================================
Public Function WriteTxtFileFrArray( _
	ByVal sTrgtFilePath, _
	ByRef cFileContents _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTxtFile
	Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
	
	Dim sFileLine
	For Each sFileLine In cFileContents
		objTxtFile.WriteLine sFileLine
	Next
	
	objTxtFile.Close
End Function
'	Call Test_WriteTxtFileFrArray()
	Private Sub Test_WriteTxtFileFrArray()
		Dim cFileContents
		Set cFileContents = CreateObject("System.Collections.ArrayList")
		cFileContents.Add "a"
		cFileContents.Add "b"
		cFileContents.Insert 1, "c"
		DIm sTrgtFilePath
		sTrgtFilePath = "C:\codes\vbs\��������CSV���`�c�[��\Test.csv"
		call WriteTxtFileFrArray( sTrgtFilePath, cFileContents )
	End Sub
