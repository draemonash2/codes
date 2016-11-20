Option Explicit

'lFileListType�j0�F�����A1:�t�@�C���A2:�t�H���_�A����ȊO�F�i�[���Ȃ�
Public Function GetFileList( _
	ByVal sTrgtDir, _
	ByRef asFileList, _
	ByVal lFileListType _
)
	Dim objFileSys
	Dim objFolder
	Dim objSubFolder
	Dim objFile
	Dim bExecStore
	Dim lLastIdx
	
	Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFileSys.GetFolder( sTrgtDir )
	
	'*** �t�H���_�p�X�i�[ ***
	Select Case lFileListType
		Case 0:    bExecStore = True
		Case 1:    bExecStore = False
		Case 2:    bExecStore = True
		Case Else: bExecStore = False
	End Select
	If bExecStore = True Then
		lLastIdx = UBound( asFileList ) + 1
		ReDim Preserve asFileList( lLastIdx )
		asFileList( lLastIdx ) = objFolder
	Else
		'Do Nothing
	End If
	
	'�t�H���_���̃T�u�t�H���_���
	'�i�T�u�t�H���_���Ȃ���΃��[�v���͒ʂ�Ȃ��j
	For Each objSubFolder In objFolder.SubFolders
		Call GetFileList( objSubFolder, asFileList, lFileListType)
	Next
	
	'*** �t�@�C���p�X�i�[ ***
	For Each objFile In objFolder.Files
		Select Case lFileListType
			Case 0:    bExecStore = True
			Case 1:    bExecStore = True
			Case 2:    bExecStore = False
			Case Else: bExecStore = False
		End Select
		If bExecStore = True Then
			'�{�X�N���v�g�t�@�C���͊i�[�ΏۊO
			If objFile.Name = WScript.ScriptName Then
				'Do Nothing
			Else
				lLastIdx = UBound( asFileList ) + 1
				ReDim Preserve asFileList( lLastIdx )
				asFileList( lLastIdx ) = objFile
			End If
		Else
			'Do Nothing
		End If
	Next
	
	Set objFolder = Nothing
	Set objFileSys = Nothing
End Function

'Dir �R�}���h�ɂ��t�@�C���ꗗ�擾�BGetFileList() ���������B
'asFileList �͔z��^�ł͂Ȃ��o���A���g�^�Ƃ��Ē�`����K�v�����邱�Ƃɒ��ӁI
'lFileListType�j0�F�����A1:�t�@�C���A2:�t�H���_�A����ȊO�F�i�[���Ȃ�
Public Function GetFileList2( _
	ByVal sTrgtDir, _
	ByRef asFileList, _
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
			asFileList = Split( sTextAll, vbNewLine )
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

'�t�H���_�����ɑ��݂��Ă���ꍇ�͉������Ȃ�
Public Function CreateDirectry( _
	ByVal sDirPath _
)
	Dim sParentDir
	Dim oFileSys
	
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	
	sParentDir = oFileSys.GetParentFolderName(sDirPath)
	
	'�e�f�B���N�g�������݂��Ȃ��ꍇ�A�ċA�Ăяo��
	If oFileSys.FolderExists( sParentDir ) = False Then
		Call CreateDirectry( sParentDir )
	End If
	
	'�f�B���N�g���쐬
	If oFileSys.FolderExists( sDirPath ) = False Then
		oFileSys.CreateFolder sDirPath
	End If
	
	Set oFileSys = Nothing
End Function

'�߂�l�j1�F�t�@�C���A2�A�t�H���_�[�A0�F�G���[�i���݂��Ȃ��p�X�j
Public Function GetFileOrFolder( _
	ByVal sChkTrgtPath _
)
	Dim oFileSys
	Dim bFolderExists
	Dim bFileExists
	
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	bFolderExists = oFileSys.FolderExists(sChkTrgtPath)
	bFileExists = oFileSys.FileExists(sChkTrgtPath)
	Set oFileSys = Nothing
	
	If bFolderExists = False And bFileExists = True Then
		GetFileOrFolder = 1 '�t�@�C��
	ElseIf bFolderExists = True And bFileExists = False Then
		GetFileOrFolder = 2 '�t�H���_�[
	Else
		GetFileOrFolder = 0 '�G���[�i���݂��Ȃ��p�X�j
	End If
End Function

'	Call Test()
'	Private Sub Test()
'		Dim objWshShell
'		Dim sCurDir
'		Set objWshShell = WScript.CreateObject( "WScript.Shell" )
'		sCurDir = objWshShell.CurrentDirectory
'		Call Include( sCurDir & "\Array.vbs" )
'		Call Include( sCurDir & "\iTunes.vbs" )
'		Call Include( sCurDir & "\ProgressBar.vbs" )
'		Call Include( sCurDir & "\StopWatch.vbs" )
'		Call Include( sCurDir & "\String.vbs" )
'		
'		Dim oStpWtch
'		
'		Set oStpWtch = New StopWatch
'		
'		oStpWtch.StartT
'	'	Dim asFileList()
'	'	ReDim asFileList(-1)
'	'	Call GetFileList( "Z:\300_Musics", asFileList, 0 )
'		Dim asFileList
'		Call GetFileList2( "Z:\300_Musics", asFileList, 1 )
'		oStpWtch.StopT
'		
'		MsgBox oStpWtch.ElapsedTime
'		Call OutputAllElement2LogFile(asFileList)
'	End Sub
'	Function Include( _
'		ByVal sOpenFile _
'		)
'		Dim objFSO
'		Dim objVbsFile
'	
'		Set objFSO = CreateObject("Scripting.FileSystemObject")
'		Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
'	
'		ExecuteGlobal objVbsFile.ReadAll()
'		objVbsFile.Close
'	
'		Set objVbsFile = Nothing
'		Set objFSO = Nothing
'	End Function

'�w��t�H���_�p�X�Ɋ܂܂��t�H���_���󂩔��肵�A��t�H���_�Ȃ�폜����B
Public Function DeleteEmptyFolder( _
	ByVal sTrgtDirPath _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim sTrgtParentDirPath
	'objLogFile.WriteLine "[Debug] called! " & sTrgtDirPath
	If objFSO.FolderExists( sTrgtDirPath ) Then
		Dim objFolder
		Set objFolder = objFSO.GetFolder( sTrgtDirPath )
		
		Dim bIsFileFolderExists
		bIsFileFolderExists = False
		
		'�T�u�t�H���_����
		Dim objSubFolder
		For Each objSubFolder In objFolder.SubFolders
			bIsFileFolderExists = True
		Next
		
		'�T�u�t�@�C������
		Dim objFile
		For Each objFile In objFolder.Files
			bIsFileFolderExists = True
		Next
		
		'objLogFile.WriteLine "[Debug] " & bIsFileFolderExists & " : " & sTrgtDirPath
		If bIsFileFolderExists = True Then
			'Do Nothing
		Else
			objFolder.Delete
			sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtDirPath )
			Call DeleteEmptyFolder( sTrgtParentDirPath )
		End If
		DeleteEmptyFolder = True
	Else
		sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtDirPath )
		Call DeleteEmptyFolder( sTrgtParentDirPath )
		DeleteEmptyFolder = False
	End If
	Set objFSO = Nothing
End Function
	Private Sub Test_DeleteEmptyFolder()
		Dim sOutStr
		sOutStr = ""
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\a\e" )
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\b.txt" )
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\c.txt" )
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\c" )
		MsgBox sOutStr
	End Sub
'	Call Test_DeleteEmptyFolder()
