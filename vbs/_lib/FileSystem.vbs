Option Explicit

'*********************************************************************
'* �O���[�o���֐���`
'*********************************************************************
' ==================================================================
' = �T�v	�t�@�C��/�t�H���_�p�X�ꗗ���擾����
' = ����	sTrgtDir		String		[in]	�Ώۃt�H���_
' = ����	asFileList		String()	[out]	�t�@�C��/�t�H���_�p�X�ꗗ
' = ����	lFileListType	Long		[in]	�擾����ꗗ�̌`��
' =													0�F����
' =													1:�t�@�C��
' =													2:�t�H���_
' =													����ȊO�F�i�[���Ȃ�
' = �ߒl	�Ȃ�
' = �o��	�Ȃ�
' ==================================================================
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
'	Call Test_GetFileList()
	Private Sub Test_GetFileList()
		Dim objWshShell
		Dim sCurDir
		Set objWshShell = WScript.CreateObject( "WScript.Shell" )
		sCurDir = objWshShell.CurrentDirectory
		Call FileSysem_Include( sCurDir & "\Array.vbs" )
		Call FileSysem_Include( sCurDir & "\iTunes.vbs" )
		Call FileSysem_Include( sCurDir & "\ProgressBar.vbs" )
		Call FileSysem_Include( sCurDir & "\StopWatch.vbs" )
		Call FileSysem_Include( sCurDir & "\String.vbs" )
		
		Dim oStpWtch
		
		Set oStpWtch = New StopWatch
		
		oStpWtch.StartT
		Dim asFileList()
		Redim Preserve asFileList(-1)
		Call GetFileList( "Z:\300_Musics", asFileList, 1 )
		oStpWtch.StopT
		
		MsgBox oStpWtch.ElapsedTime
		Call OutputAllElement2LogFile(asFileList)
	End Sub

' ==================================================================
' = �T�v	�t�@�C��/�t�H���_�p�X�ꗗ���擾����
' = ����	sTrgtDir		String		[in]	�Ώۃt�H���_
' = ����	asFileList		Variant		[out]	�t�@�C��/�t�H���_�p�X�ꗗ
' = ����	lFileListType	Long		[in]	�擾����ꗗ�̌`��
' =													0�F����
' =													1:�t�@�C��
' =													2:�t�H���_
' =													����ȊO�F�i�[���Ȃ�
' = �ߒl	�Ȃ�
' = �o��	�EDir �R�}���h�ɂ��t�@�C���ꗗ�擾�BGetFileList() ���������B
' =			�EasFileList �͔z��^�ł͂Ȃ��o���A���g�^�Ƃ��Ē�`����
' =			  �K�v�����邱�Ƃɒ��ӁI
' ==================================================================
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
'	Call Test_GetFileList2()
	Private Sub Test_GetFileList2()
		Dim objWshShell
		Dim sCurDir
		Set objWshShell = WScript.CreateObject( "WScript.Shell" )
		sCurDir = objWshShell.CurrentDirectory
		Call FileSysem_Include( sCurDir & "\Array.vbs" )
		Call FileSysem_Include( sCurDir & "\iTunes.vbs" )
		Call FileSysem_Include( sCurDir & "\ProgressBar.vbs" )
		Call FileSysem_Include( sCurDir & "\StopWatch.vbs" )
		Call FileSysem_Include( sCurDir & "\String.vbs" )
		
		Dim oStpWtch
		
		Set oStpWtch = New StopWatch
		
		oStpWtch.StartT
		Dim asFileList
		Call GetFileList2( "Z:\300_Musics", asFileList, 1 )
		oStpWtch.StopT
		
		MsgBox oStpWtch.ElapsedTime
		Call OutputAllElement2LogFile(asFileList)
	End Sub

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
' =			�EArray�R���N�V�����Ɋi�[����
' ==================================================================
Public Function GetFileList3( _
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
'	Call Test_GetFileList3()
	Private Sub Test_GetFileList3()
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Dim sCurDir
		sCurDir = objFSO.GetParentFolderName( WScript.ScriptFullName )
		
		msgbox sCurDir
		
		Dim cFileList
		Set cFileList = CreateObject("System.Collections.ArrayList")
		Call GetFileList3( sCurDir, cFileList, 1 )
		
		dim sFilePath
		dim sOutput
		sOutput = ""
		for each sFilePath in cFileList
			sOutput = sOutput & vbNewLine & sFilePath
		next
		MsgBox sOutput
	End Sub

' ==================================================================
' = �T�v	�t�H���_���쐬����
' = ����	sDirPath		String		[in]	�쐬�Ώۃt�H���_
' = �ߒl	�Ȃ�
' = �o��	�E�쐬�Ώۃt�H���_�̐e�f�B���N�g�������݂��Ȃ��ꍇ�A
' =			  �ċA�I�ɐe�t�H���_���쐬����
' =			�E�t�H���_�����ɑ��݂��Ă���ꍇ�͉������Ȃ�
' ==================================================================
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

' ==================================================================
' = �T�v	�t�@�C�����t�H���_���𔻒肷��
' = ����	sChkTrgtPath	String		[in]	�`�F�b�N�Ώۃt�H���_
' = �ߒl					Long				���茋��
' =													1) �t�@�C��
' =													2) �t�H���_�[
' =													0) �G���[�i���݂��Ȃ��p�X�j
' = �o��	FileSystemObject ���g���Ă���̂ŁA�t�@�C��/�t�H���_��
' =			���݊m�F�ɂ��g�p�\�B
' ==================================================================
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

' ==================================================================
' = �T�v	�w��t�H���_�p�X�Ɋ܂܂��t�H���_���󂩔��肵�A
' =			��t�H���_�Ȃ�폜����B
' = ����	sTrgtPath	String		[in]	�`�F�b�N�Ώۃt�H���_
' = �ߒl				String				�폜���ʃ��O
' = �o��	�Ȃ�
' ==================================================================
Public Function DeleteEmptyFolder( _
	ByVal sTrgtPath _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim sTrgtParentDirPath
	Dim sRetStr
	'objLogFile.WriteLine "[Debug] called! " & sTrgtPath
	If objFSO.FolderExists( sTrgtPath ) Then
		Dim objFolder
		Set objFolder = objFSO.GetFolder( sTrgtPath )
		
		Dim bIsFileFolderExists
		bIsFileFolderExists = False
		
		'�T�u�t�H���_����
		Dim objSubFolder
		For Each objSubFolder In objFolder.SubFolders
			bIsFileFolderExists = True
			Exit For
		Next
		
		'�T�u�t�@�C������
		Dim objFile
		For Each objFile In objFolder.Files
			bIsFileFolderExists = True
			Exit For
		Next
		
		'objLogFile.WriteLine "[Debug] " & bIsFileFolderExists & " : " & sTrgtPath
		If bIsFileFolderExists = True Then
			sRetStr = sRetStr & vbNewLine & "[Folder] exists / stay   / -- / " & sTrgtPath
		Else
			objFolder.Delete
			sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtPath )
			sRetStr = sRetStr & vbNewLine & "[Folder] empty  / delete / �� / " & sTrgtPath
			sRetStr = sRetStr & DeleteEmptyFolder( sTrgtParentDirPath )
		End If
	ElseIf objFSO.FileExists( sTrgtPath ) Then
		sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtPath )
		sRetStr = sRetStr & vbNewLine & "[File	]		 / stay   / �� / " & sTrgtPath
		sRetStr = sRetStr & DeleteEmptyFolder( sTrgtParentDirPath )
	Else
		sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtPath )
		sRetStr = sRetStr & vbNewLine & "[-		]		 / stay   / �� / " & sTrgtPath
		sRetStr = sRetStr & DeleteEmptyFolder( sTrgtParentDirPath )
	End If
	DeleteEmptyFolder = sRetStr
	Set objFSO = Nothing
End Function
'	Call Test_DeleteEmptyFolder()
	Private Sub Test_DeleteEmptyFolder()
		Dim sOutStr
		sOutStr = ""
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\a\e\e.txt" )
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\a\e" )
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\b.txt" )
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\c.txt" )
		sOutStr = sOutStr & vbNewLine & DeleteEmptyFolder( "C:\codes\vbs\test\c" )
		MsgBox sOutStr
	End Sub

' ==================================================================
' = �T�v	�w��p�X�����݂���ꍇ�A"_XXX" ��t�^���ĕԋp����
' = ����	sTrgtPath		String		[in]	�Ώۃp�X
' = ����	sAddedPath		String		[out]	�t�^��̃p�X
' = ����	lAddedPathType	Long		[out]	�t�^��̃p�X���
' =												  1: �t�@�C��
' =												  2: �t�H���_
' = �ߒl					Boolean				�擾����
' = �o��	�{�֐��ł́A�t�@�C��/�t�H���_�͍쐬���Ȃ��B
' ==================================================================
Public Function GetNotExistPath( _
	ByVal sTrgtPath, _
	ByRef sAddedPath, _
	ByRef lAddedPathType _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim bFolderExists
	Dim bFileExists
	bFolderExists = objFSO.FolderExists( sTrgtPath )
	bFileExists = objFSO.FileExists( sTrgtPath )
	
	If bFolderExists = False And bFileExists = True Then
		sAddedPath = GetFileNotExistPath( sTrgtPath )
		lAddedPathType = 1
		GetNotExistPath = True
	ElseIf bFolderExists = True And bFileExists = False Then
		sAddedPath = GetFolderNotExistPath( sTrgtPath )
		lAddedPathType = 2
		GetNotExistPath = True
	Else
		sAddedPath = sTrgtPath
		lAddedPathType = 0
		GetNotExistPath = False
	End If
End Function
	'Call Test_GetNotExistPath()
	Private Sub Test_GetNotExistPath()
		Dim sOutStr
		Dim sAddedPath
		Dim lAddedPathType
		Dim bRet
																						   sOutStr = ""
																						   sOutStr = sOutStr & vbNewLine & "*** test start! ***"
		bRet = GetNotExistPath("C:\codes\vba", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
		bRet = GetNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas", sAddedPath, lAddedPathType): sOutStr = sOutStr & vbNewLine & bRet & " / " & lAddedPathType & " : " & sAddedPath
																						   sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
		MsgBox sOutStr
	End Sub

' ==================================================================
' = �T�v	�t�H���_�I���_�C�A���O��\������
' = ����	sInitPath	String	[in]  �f�t�H���g�t�H���_�p�X
' = �ߒl				String		  �t�H���_�I������
' = �o��	�E���݂��Ȃ��t�H���_�p�X��I�������ꍇ�A�󕶎����ԋp����
' =			�E�L�����Z�������������ꍇ�A�󕶎����ԋp����
' ==================================================================
Private Function ShowFolderSelectDialog( _
	ByVal sInitPath _
)
	Const msoFileDialogFolderPicker = 4
	Const xlMinimized = -4140
	
	Dim objExcel
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False '��\���ɂ��Ă�����ۂɂ�����ƕ\�����ꂿ�Ⴄ�B
	objExcel.WindowState = xlMinimized '��L���R����ŏ��������Ƃ��B
	
	Dim fdDialog
	Set fdDialog = objExcel.FileDialog(msoFileDialogFolderPicker)
	fdDialog.Title = "�t�H���_��I�����Ă��������i�󗓂̏ꍇ�͐e�t�H���_���I������܂��j"
	If sInitPath = "" Then
		'Do Nothing
	Else
		If Right(sInitPath, 1) = "\" Then
			fdDialog.InitialFileName = sInitPath
		Else
			fdDialog.InitialFileName = sInitPath & "\"
		End If
	End If
	
	'�_�C�A���O�\��
	Dim lResult
	lResult = fdDialog.Show()
	If lResult <> -1 Then '�L�����Z������
		ShowFolderSelectDialog = ""
	Else
		Dim sSelectedPath
		sSelectedPath = fdDialog.SelectedItems.Item(1)
		If CreateObject("Scripting.FileSystemObject").FolderExists( sSelectedPath ) Then
			ShowFolderSelectDialog = sSelectedPath
		Else
			ShowFolderSelectDialog = ""
		End If
	End If
	
	Set fdDialog = Nothing
End Function
'	Call Test_ShowFolderSelectDialog()
	Private Sub Test_ShowFolderSelectDialog()
		Dim objWshShell
		Set objWshShell = CreateObject("WScript.Shell")
		
		Dim sInitPath
		sInitPath = objWshShell.SpecialFolders("Desktop")
		'sInitPath = ""
		
		MsgBox ShowFolderSelectDialog( sInitPath )
	End Sub

' ==================================================================
' = �T�v	�t�@�C���i�P��j�I���_�C�A���O��\������
' = ����	sInitPath	String	[in]  �f�t�H���g�t�@�C���p�X
' = ����	sFilters	String	[in]  �I�����̃t�B���^(��)
' = �ߒl				String		  �t�@�C���I������
' = �o��	(��)�_�C�A���O�̃t�B���^�w����@�͈ȉ��B
' =				 ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =					   �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =					   �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =					   �E�t�B���^����������ꍇ�A","�ŋ�؂�
' =			sFilters ���ȗ��������͋󕶎��̏ꍇ�A�t�B���^���N���A����B
' ==================================================================
Private Function ShowFileSelectDialog( _
	ByVal sInitPath, _
	ByVal sFilters _
)
	Const msoFileDialogFilePicker = 3
	Const xlMinimized = -4140
	
	Dim objExcel
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False '��\���ɂ��Ă�����ۂɂ�����ƕ\�����ꂿ�Ⴄ�B
	objExcel.WindowState = xlMinimized '��L���R����ŏ��������Ƃ��B
	
	Dim fdDialog
	Set fdDialog = objExcel.FileDialog(msoFileDialogFilePicker)
	fdDialog.Title = "�t�@�C����I�����Ă�������"
	fdDialog.AllowMultiSelect = False
	If sInitPath = "" Then
		'Do Nothing
	Else
		fdDialog.InitialFileName = sInitPath
	End If
	Call SetDialogFilters(sFilters, fdDialog) '�t�B���^�ǉ�
	
	'�_�C�A���O�\��
	Dim lResult
	lResult = fdDialog.Show()
	If lResult <> -1 Then '�L�����Z������
		ShowFileSelectDialog = ""
	Else
		Dim sSelectedPath
		sSelectedPath = fdDialog.SelectedItems.Item(1)
		If CreateObject("Scripting.FileSystemObject").FileExists( sSelectedPath ) Then
			ShowFileSelectDialog = sSelectedPath
		Else
			ShowFileSelectDialog = ""
		End If
	End If
	
	Set fdDialog = Nothing
End Function
'	Call Test_ShowFileSelectDialog()
	Private Sub Test_ShowFileSelectDialog()
		Dim objWshShell
		Set objWshShell = CreateObject("WScript.Shell")
		
		Dim sInitPath
		sInitPath = objWshShell.SpecialFolders("Desktop") & "\test.txt"
		'sInitPath = ""
		
		Dim sFilters
		'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png"
		'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv"
		'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png,�e�L�X�g�t�@�C��/*.txt; *.csv"
		sFilters = ""
		
		MsgBox ShowFileSelectDialog( sInitPath, sFilters )
	End Sub

' ==================================================================
' = �T�v	�t�@�C���i�����j�I���_�C�A���O��\������
' = ����	asSelectedFiles String()	[out] �I�����ꂽ�t�@�C���p�X�ꗗ
' = ����	sInitPath		String		[in]  �f�t�H���g�t�@�C���p�X
' = ����	sFilters		String		[in]  �I�����̃t�B���^(��)
' = �ߒl	�Ȃ�
' = �o��	(��)�_�C�A���O�̃t�B���^�w����@�͈ȉ��B
' =				 ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =					   �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =					   �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =					   �E�t�B���^����������ꍇ�A","�ŋ�؂�
' =			sFilters ���ȗ��������͋󕶎��̏ꍇ�A�t�B���^���N���A����B
' ==================================================================
Private Function ShowFilesSelectDialog( _
	ByRef asSelectedFiles(), _
	ByVal sInitPath, _
	ByVal sFilters _
)
	Const msoFileDialogFilePicker = 3
	Const xlMinimized = -4140
	
	Dim objExcel
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False '��\���ɂ��Ă�����ۂɂ�����ƕ\�����ꂿ�Ⴄ�B
	objExcel.WindowState = xlMinimized '��L���R����ŏ��������Ƃ��B
	
	Dim fdDialog
	Set fdDialog = objExcel.FileDialog(msoFileDialogFilePicker)
	fdDialog.Title = "�t�@�C����I�����Ă��������i�����j"
	fdDialog.AllowMultiSelect = True
	If sInitPath = "" Then
		'Do Nothing
	Else
		fdDialog.InitialFileName = sInitPath
	End If
	Call SetDialogFilters(sFilters, fdDialog) '�t�B���^�ǉ�
	
	'�_�C�A���O�\��
	Dim lResult
	lResult = fdDialog.Show()
	If lResult <> -1 Then '�L�����Z������
		ReDim Preserve asSelectedFiles(0)
		asSelectedFiles(0) = ""
	Else
		Dim lSelNum
		lSelNum = fdDialog.SelectedItems.Count
		ReDim Preserve asSelectedFiles(lSelNum - 1)
		Dim lSelIdx
		For lSelIdx = 0 To lSelNum - 1
			asSelectedFiles(lSelIdx) = fdDialog.SelectedItems(lSelIdx + 1)
		Next
	End If
	
	Set fdDialog = Nothing
End Function
'	Call Test_ShowFilesSelectDialog()
	Private Sub Test_ShowFilesSelectDialog()
		Dim objWshShell
		Set objWshShell = CreateObject("WScript.Shell")
		
		Dim sFilters
		'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png"
		'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv"
		'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png,�e�L�X�g�t�@�C��/*.txt; *.csv"
		sFilters = "�S�Ẵt�@�C��/*.*,�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png,�e�L�X�g�t�@�C��/*.txt; *.csv"
		
		Dim sInitPath
		'sInitPath = objWshShell.SpecialFolders("Desktop") & "\test.txt"
		sInitPath = ""
		
		Dim asSelectedFiles()
		Call ShowFilesSelectDialog( _
					asSelectedFiles, _
					sInitPath, _
					sFilters _
				)
		Dim sBuf
		sBuf = ""
		sBuf = sBuf & vbNewLine & UBound(asSelectedFiles) + 1
		Dim lSelIdx
		For lSelIdx = 0 To UBound(asSelectedFiles)
			sBuf = sBuf & vbNewLine & asSelectedFiles(lSelIdx)
		Next
		MsgBox sBuf
	End Sub

' ==================================================================
' = �T�v	�h���C�u���擾�i�h���C�u���^�[�w��j
' = ����	sDriveLetter	String		[in]	�h���C�u���^�[
' = ����	lGetInfoType	Long		[in]	�擾�����
' =													1) �{�����[�����x��
' =													2) �t�H���_
' =													3) ���[�g�t�H���_
' =													4) ���
' =													5) �t�@�C���V�X�e��
' =													6) �e��
' =													7) �󂫗̈�
' =													8) �V���A���i���o�[
' = ����	sDriveInfo		String		[out]	�h���C�u���
' = �ߒl					Boolean				�擾����
' = �o��	�E�l�b�g���[�N�h���C�u�������\
' ==================================================================
Public Function GetDriveInfoFromDriveLetter( _
	ByVal sDriveLetter, _
	ByVal lGetInfoType, _
	ByRef sDriveInfo _
)
	Dim lDrvLtrIdx
	Dim lDrvLtrAscStrt
	Dim lDrvLtrAscLast
	lDrvLtrIdx = asc(sDriveLetter)
	lDrvLtrAscStrt = asc("A")
	lDrvLtrAscLast = asc("Z")
	If lDrvLtrIdx >= lDrvLtrAscStrt And lDrvLtrIdx <= lDrvLtrAscLast Then
		'Do Nothing
	Else
		GetDriveInfoFromDriveLetter = False
		Exit Function
	End If
	
	sDriveInfo = ""
	Dim DRIVE_TYPE_TABLE
	DRIVE_TYPE_TABLE = Array("Unknown", "Removable", "HDD", "Network", "CD-ROM", "RAM")
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If objFSO.DriveExists(sDriveLetter) Then
		Dim objDrive
		Set objDrive = objFSO.GetDrive(sDriveLetter)
		If objDrive.IsReady = True Then
			GetDriveInfoFromDriveLetter = True
			Select Case lGetInfoType
				Case 1:		sDriveInfo = objDrive.VolumeName					'�{�����[�����x��
				Case 2:		sDriveInfo = objDrive.Path							'�t�H���_
				Case 3:		sDriveInfo = objDrive.RootFolder					'���[�g�t�H���_
				Case 4:		sDriveInfo = DRIVE_TYPE_TABLE(objDrive.DriveType)	'���
				Case 5:		sDriveInfo = objDrive.FileSystem					'�t�@�C���V�X�e��
				Case 6:		sDriveInfo = FormatNumber(objDrive.TotalSize, 0)	'�e��
				Case 7:		sDriveInfo = FormatNumber(objDrive.FreeSpace, 0)	'�󂫗̈�
				Case 8:		sDriveInfo = Hex(objDrive.SerialNumber)				'�V���A���i���o�[
				Case Else:	GetDriveInfoFromDriveLetter = False
			End Select
		Else
			GetDriveInfoFromDriveLetter = False
		End If
	Else
		GetDriveInfoFromDriveLetter = False
	End If
End Function
'	Call Test_GetDriveInfoFromDriveLetter()
	Private Sub Test_GetDriveInfoFromDriveLetter()
		Dim sBuf
		Dim bRet
		Dim sDriveInfo
		sBuf = ""
																 sBuf = sBuf & vbNewLine &		  "*** C �h���C�u ***"
		bRet = GetDriveInfoFromDriveLetter("C", 1, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �{�����[�����x���F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 2, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�H���_�F"			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ���[�g�t�H���_�F"		& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 4, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ��ށF"				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 5, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���V�X�e���F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 6, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �e�ʁF"				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 7, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �󂫗̈�F"			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 8, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �V���A���i���o�[�F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 9, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  "						& sDriveInfo
																 sBuf = sBuf & vbNewLine &		  "*** X �h���C�u ***"
		bRet = GetDriveInfoFromDriveLetter("X", 1, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �{�����[�����x���F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 2, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�H���_�F"			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ���[�g�t�H���_�F"		& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 4, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ��ށF"				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 5, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���V�X�e���F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 6, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �e�ʁF"				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 7, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �󂫗̈�F"			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 8, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �V���A���i���o�[�F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 9, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  "						& sDriveInfo
																 sBuf = sBuf & vbNewLine &		  "*** Z �h���C�u ***"
		bRet = GetDriveInfoFromDriveLetter("Z", 1, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �{�����[�����x���F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 2, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�H���_�F"			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ���[�g�t�H���_�F"		& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 4, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ��ށF"				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 5, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���V�X�e���F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 6, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �e�ʁF"				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 7, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �󂫗̈�F"			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 8, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  �V���A���i���o�[�F"	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 9, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  "						& sDriveInfo
																 sBuf = sBuf & vbNewLine &		  "*** E �h���C�u ***"
		bRet = GetDriveInfoFromDriveLetter("E", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ���[�g�t�H���_�F"		& sDriveInfo
		MsgBox sBuf
	End Sub

' ==================================================================
' = �T�v	�h���C�u���擾�i�{�����[�����x���w��j
' = ����	sVolumeLabel	String		[in]	�{�����[�����x��
' = ����	lGetInfoType	Long		[in]	�擾�����
' =													1) �{�����[�����x��
' =													2) �t�H���_
' =													3) ���[�g�t�H���_
' =													4) ���
' =													5) �t�@�C���V�X�e��
' =													6) �e��
' =													7) �󂫗̈�
' =													8) �V���A���i���o�[
' = ����	sDriveInfo		String		[out]	�h���C�u���
' = �ߒl					Boolean				�擾����
' = �o��	�EsVolumeLabel ���󕶎���̏ꍇ�A�{�����[�����x�����ݒ�
' =			  ����Ă��Ȃ��h���C�u�̏�񂪕ԋp�����
' =			�E�l�b�g���[�N�h���C�u�������\
' ==================================================================
Public Function GetDriveInfoFromVolumeLabel( _
	ByVal sVolumeLabel, _
	ByVal lGetInfoType, _
	ByRef sDriveInfo _
)
	Dim lDrvLtrAscStrt
	Dim lDrvLtrAscLast
	lDrvLtrAscStrt = asc("A")
	lDrvLtrAscLast = asc("Z")
	
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Dim DRIVE_TYPE_TABLE
	DRIVE_TYPE_TABLE = Array("Unknown", "Removable", "HDD", "Network", "CD-ROM", "RAM")
	Dim bIsContinue
	bIsContinue = False
	sDriveInfo = ""
	GetDriveInfoFromVolumeLabel = False
	
	On Error Resume Next
	Dim lDrvLtrIdx
	For lDrvLtrIdx = lDrvLtrAscStrt to lDrvLtrAscLast
		Dim sDriveLetter
		sDriveLetter = Chr(lDrvLtrIdx)
		If Err.Number = 0 Then
			If objFSO.DriveExists(sDriveLetter) Then
				Dim objDrive
				Set objDrive = objFSO.GetDrive(sDriveLetter)
				If objDrive.VolumeName = sVolumeLabel Then
					If objDrive.IsReady = True Then
						GetDriveInfoFromVolumeLabel = True
						Select Case lGetInfoType
							Case 1:		sDriveInfo = objDrive.VolumeName					'�{�����[�����x��
							Case 2:		sDriveInfo = objDrive.Path							'�t�H���_
							Case 3:		sDriveInfo = objDrive.RootFolder					'���[�g�t�H���_
							Case 4:		sDriveInfo = DRIVE_TYPE_TABLE(objDrive.DriveType)	'���
							Case 5:		sDriveInfo = objDrive.FileSystem					'�t�@�C���V�X�e��
							Case 6:		sDriveInfo = FormatNumber(objDrive.TotalSize, 0)	'�e��
							Case 7:		sDriveInfo = FormatNumber(objDrive.FreeSpace, 0)	'�󂫗̈�
							Case 8:		sDriveInfo = Hex(objDrive.SerialNumber)				'�V���A���i���o�[
							Case Else:	GetDriveInfoFromVolumeLabel = False
						End Select
						bIsContinue = False
					Else
						bIsContinue = False
					End If
				Else
					bIsContinue = True
				End If
			Else
				bIsContinue = True
			End If
		Else
			bIsContinue = False
		End If
		If bIsContinue = True Then
			'Do Nothing
		Else
			Exit For
		End If
	Next
	On Error Goto 0
End Function
'	Call Test_GetDriveInfoFromVolumeLabel()
	Private Sub Test_GetDriveInfoFromVolumeLabel()
		Dim sBuf
		Dim bRet
		Dim sDriveInfo
		sBuf = ""
																			sBuf = sBuf & vbNewLine &		 "*** �h���C�u�� SD256G ***"
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 1, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	�{�����[�����x���F" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 2, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	�t�H���_�F"			& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 3, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	���[�g�t�H���_�F"	& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 4, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	��ށF"				& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 5, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���V�X�e���F" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 6, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	�e�ʁF"				& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 7, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	�󂫗̈�F"			& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 8, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	�V���A���i���o�[�F" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 9, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	"					& sDriveInfo
																			sBuf = sBuf & vbNewLine &		 "*** �h���C�u�� logitechdd3t ***"
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 1, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	�{�����[�����x���F" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 2, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�H���_�F"			& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	���[�g�t�H���_�F"	& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 4, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	��ށF"				& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 5, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���V�X�e���F" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 6, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	�e�ʁF"				& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 7, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	�󂫗̈�F"			& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 8, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	�V���A���i���o�[�F" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 9, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	"					& sDriveInfo
																			sBuf = sBuf & vbNewLine &		 "*** �h���C�u�� - ***"
		bRet = GetDriveInfoFromVolumeLabel("-", 3, sDriveInfo) :			sBuf = sBuf & vbNewLine & bRet & "	���[�g�t�H���_�F"	& sDriveInfo
																			sBuf = sBuf & vbNewLine &		 "*** �h���C�u�� """" ***"
		bRet = GetDriveInfoFromVolumeLabel("", 3, sDriveInfo) :				sBuf = sBuf & vbNewLine & bRet & "	���[�g�t�H���_�F"	& sDriveInfo
		MsgBox sBuf
	End Sub

' ==================================================================
' = �T�v	�t�@�C�����擾
' = ����	sTrgtPath		String		[in]	�t�@�C���p�X
' = ����	lGetInfoType	Long		[in]	�擾����� (��1)
' = ����	vFileInfo		Variant		[out]	�t�@�C����� (��1)
' = �ߒl					Boolean				�擾����
' = �o��	�ȉ��A�Q�ƁB
' =		(��1) �t�@�C�����
' =			[����]	[����]					[�v���p�e�B��]		[�f�[�^�^]				[Get/Set]	[�o�͗�]
' =			1		�t�@�C����				Name				vbString	������^	Get/Set		03 Ride Featuring Tony Matterhorn.MP3
' =			2		�t�@�C���T�C�Y			Size				vbLong		�������^	Get			4286923
' =			3		�t�@�C�����			Type				vbString	������^	Get			MPEG layer 3
' =			4		�t�@�C���i�[��h���C�u	Drive				vbString	������^	Get			Z:
' =			5		�t�@�C���p�X			Path				vbString	������^	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =			6		�e�t�H���_				ParentFolder		vbString	������^	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =			7		MS-DOS�`���t�@�C����	ShortName			vbString	������^	Get			03 Ride Featuring Tony Matterhorn.MP3
' =			8		MS-DOS�`���p�X			ShortPath			vbString	������^	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =			9		�쐬����				DateCreated			vbDate		���t�^		Get			2015/08/19 0:54:45
' =			10		�A�N�Z�X����			DateLastAccessed	vbDate		���t�^		Get			2016/10/14 6:00:30
' =			11		�X�V����				DateLastModified	vbDate		���t�^		Get			2016/10/14 6:00:30
' =			12		����					Attributes			vbLong		�������^	(��2)		32
' =		(��2) ����
' =			[�l]				[����]										[������]	[Get/Set]
' =			1  �i0b00000001�j	�ǂݎ���p�t�@�C��						ReadOnly	Get/Set
' =			2  �i0b00000010�j	�B���t�@�C��								Hidden		Get/Set
' =			4  �i0b00000100�j	�V�X�e���E�t�@�C��							System		Get/Set
' =			8  �i0b00001000�j	�f�B�X�N�h���C�u�E�{�����[���E���x��		Volume		Get
' =			16 �i0b00010000�j	�t�H���_�^�f�B���N�g��						Directory	Get
' =			32 �i0b00100000�j	�O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1	Archive		Get/Set
' =			64 �i0b01000000�j	�����N�^�V���[�g�J�b�g						Alias		Get
' =			128�i0b10000000�j	���k�t�@�C��								Compressed	Get
' ==================================================================
Public Function GetFileInfo( _
	ByVal sTrgtPath, _
	ByVal lGetInfoType, _
	ByRef vFileInfo _
)
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists( sTrgtPath ) Then
		'Do Nothing
	Else
		vFileInfo = ""
		GetFileInfo = False
		Exit Function
	End If
	
	Dim objFile
	Set objFile = objFSO.GetFile(sTrgtPath)
	
	vFileInfo = ""
	GetFileInfo = True
	Select Case lGetInfoType
		Case 1:		vFileInfo = objFile.Name				'�t�@�C����
		Case 2:		vFileInfo = objFile.Size				'�t�@�C���T�C�Y
		Case 3:		vFileInfo = objFile.Type				'�t�@�C�����
		Case 4:		vFileInfo = objFile.Drive				'�t�@�C���i�[��h���C�u
		Case 5:		vFileInfo = objFile.Path				'�t�@�C���p�X
		Case 6:		vFileInfo = objFile.ParentFolder		'�e�t�H���_
		Case 7:		vFileInfo = objFile.ShortName			'MS-DOS�`���t�@�C����
		Case 8:		vFileInfo = objFile.ShortPath			'MS-DOS�`���p�X
		Case 9:		vFileInfo = objFile.DateCreated			'�쐬����
		Case 10:	vFileInfo = objFile.DateLastAccessed	'�A�N�Z�X����
		Case 11:	vFileInfo = objFile.DateLastModified	'�X�V����
		Case 12:	vFileInfo = objFile.Attributes			'����
		Case Else:	GetFileInfo = False
	End Select
End Function
'	Call Test_GetFileInfo()
	Private Sub Test_GetFileInfo()
		Dim sBuf
		Dim bRet
		Dim vFileInfo
		sBuf = ""
		Dim sTrgtPath
		sTrgtPath = "C:\codes\vbs\_lib\FileSystem.vbs"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFileInfo( sTrgtPath,	1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C�����F"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	2, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���T�C�Y�F"		  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	3, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C����ށF"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	4, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���i�[��h���C�u�F" & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	5, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���p�X�F"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	6, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�e�t�H���_�F"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	7, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS�`���t�@�C�����F"   & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	8, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS�`���p�X�F"		  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	9, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�쐬�����F"				  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 10, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�A�N�Z�X�����F"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 11, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�X�V�����F"				  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 12, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�����F"					  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 13, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�F"						  & vFileInfo
		sTrgtPath = "C:\codes\vbs\_lib\dummy.vbs"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFileInfo( sTrgtPath,	1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C�����F"			  & vFileInfo
		MsgBox sBuf
	End Sub

'�t�@�C�����́u�t�@�C�����v�u�����v���ݒ�\
'�������A�ȉ��̃��\�b�h�ɂĕύX�\�Ȃ��߁A�������Ȃ�
'  �t�@�C�����F objFSO.MoveFile
'  �����F objFSO.GetFile( "C:\codes\a.txt" ).Attributes
Public Function SetFileInfo( _
   ByVal sTrgtPath, _
   ByVal lSetInfoType, _
   ByVal vFileInfo _
)
	'Do Nothing
End Function

' ==================================================================
' = �T�v	�t�H���_���擾
' = ����	sTrgtPath		String		[in]	�t�H���_�p�X
' = ����	lGetInfoType	Long		[in]	�擾����� (��1)
' = ����	vFolderInfo		Variant		[out]	�t�H���_��� (��1)
' = �ߒl					Boolean				�擾����
' = �o��	�ȉ��A�Q�ƁB
' =		(��1) �t�H���_���
' =			[����]	[����]					[�v���p�e�B��]		[�f�[�^�^]				[Get/Set]	[�o�͗�]
' =			1		�t�H���_��				Name				vbString	������^	Get/Set		Sacrifice
' =			2		�t�H���_�T�C�Y			Size				vbLong		�������^	Get			80613775
' =			3		�t�@�C�����			Type				vbString	������^	Get			�t�@�C�� �t�H���_�[
' =			4		�t�@�C���i�[��h���C�u	Drive				vbString	������^	Get			Z:
' =			5		�t�H���_�p�X			Path				vbString	������^	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =			6		���[�g �t�H���_			IsRootFolder		vbBoolean	�u�[���^	Get			False
' =			7		MS-DOS�`���t�@�C����	ShortName			vbString	������^	Get			Sacrifice
' =			8		MS-DOS�`���p�X			ShortPath			vbString	������^	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =			9		�쐬����				DateCreated			vbDate		���t�^		Get			2015/08/19 0:54:44
' =			10		�A�N�Z�X����			DateLastAccessed	vbDate		���t�^		Get			2015/08/19 0:54:44
' =			11		�X�V����				DateLastModified	vbDate		���t�^		Get			2015/04/18 3:38:36
' =			12		����					Attributes			vbLong		�������^	(��2)		16
' =		(��2) ����
' =			[�l]				[����]										[������]	[Get/Set]
' =			1  �i0b00000001�j	�ǂݎ���p�t�@�C��						ReadOnly	Get/Set
' =			2  �i0b00000010�j	�B���t�@�C��								Hidden		Get/Set
' =			4  �i0b00000100�j	�V�X�e���E�t�@�C��							System		Get/Set
' =			8  �i0b00001000�j	�f�B�X�N�h���C�u�E�{�����[���E���x��		Volume		Get
' =			16 �i0b00010000�j	�t�H���_�^�f�B���N�g��						Directory	Get
' =			32 �i0b00100000�j	�O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1	Archive		Get/Set
' =			64 �i0b01000000�j	�����N�^�V���[�g�J�b�g						Alias		Get
' =			128�i0b10000000�j	���k�t�@�C��								Compressed	Get
' ==================================================================
Public Function GetFolderInfo( _
	ByVal sTrgtPath, _
	ByVal lGetInfoType, _
	ByRef vFolderInfo _
)
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists( sTrgtPath ) Then
		'Do Nothing
	Else
		vFolderInfo = ""
		GetFolderInfo = False
		Exit Function
	End If
	
	Dim objFolder
	Set objFolder = objFSO.GetFolder(sTrgtPath)
	
	vFolderInfo = ""
	GetFolderInfo = True
	Select Case lGetInfoType
		Case 1:		vFolderInfo = objFolder.Name				'�t�H���_��
		Case 2:		vFolderInfo = objFolder.Size				'�t�H���_�T�C�Y
		Case 3:		vFolderInfo = objFolder.Type				'�t�@�C�����
		Case 4:		vFolderInfo = objFolder.Drive				'�t�@�C���i�[��h���C�u
		Case 5:		vFolderInfo = objFolder.Path				'�t�H���_�p�X
		Case 6:		vFolderInfo = objFolder.IsRootFolder		'���[�g �t�H���_
		Case 7:		vFolderInfo = objFolder.ShortName			'MS-DOS�`���t�@�C����
		Case 8:		vFolderInfo = objFolder.ShortPath			'MS-DOS�`���p�X
		Case 9:		vFolderInfo = objFolder.DateCreated			'�쐬����
		Case 10:	vFolderInfo = objFolder.DateLastAccessed	'�A�N�Z�X����
		Case 11:	vFolderInfo = objFolder.DateLastModified	'�X�V����
		Case 12:	vFolderInfo = objFolder.Attributes			'����
		Case Else:	GetFolderInfo = False
	End Select
End Function
'	Call Test_GetFolderInfo()
	Private Sub Test_GetFolderInfo()
		Dim sBuf
		Dim bRet
		Dim vFolderInfo
		sBuf = ""
		Dim sTrgtPath
		sTrgtPath = "C:\codes\vbs\lib"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFolderInfo( sTrgtPath, 1,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C�����F"			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 2,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���T�C�Y�F"		  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 3,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C����ށF"			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 4,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���i�[��h���C�u�F" & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 5,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���p�X�F"			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 6,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�e�t�H���_�F"			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 7,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS�`���t�@�C�����F"   & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 8,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS�`���p�X�F"		  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 9,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�쐬�����F"				  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 10, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�A�N�Z�X�����F"			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 11, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�X�V�����F"				  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 12, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�����F"					  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 13, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�F"						  & vFolderInfo
		sTrgtPath = "C:\codes\vbs\libs"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFolderInfo( sTrgtPath, 1,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C�����F"			  & vFolderInfo
		MsgBox sBuf
	End Sub

'�t�H���_���́u�t�@�C�����v�u�����v���ݒ�\
'�������A�ȉ��̃��\�b�h�ɂĕύX�\�Ȃ��߁A�������Ȃ�
'  �t�@�C�����F objFSO.MoveFolder
'  �����F objFSO.GetFolder( "C:\codes" ).Attributes
Public Function SetFolderInfo( _
   ByVal sTrgtPath, _
   ByVal lSetInfoType, _
   ByVal vFolderInfo _
)
	'Do Nothing
End Function

' ==================================================================
' = �T�v	�t�@�C���ڍ׏��擾
' = ����	sTrgtPath			String		[in]	�t�@�C���p�X
' = ����	lFileInfoTagIndex	Long		[in]	�擾����ʔԍ�(��1)
' = ����	vFileInfoValue		Variant		[out]	�t�@�C���ڍ׏��
' = ����	vFileInfoTitle		Variant		[out]	�t�@�C���ڍ׏��^�C�g��
' = ����	sErrorDetail		String		[out]	�擾���ʃG���[�ڍ�(��2)
' = �ߒl						Boolean				�擾����
' = �o��	(��1) �擾�ł�����͂n�r�̃o�[�W�����ɂ���ĈقȂ�B
' =				  ���O�� GetFileDetailInfoIndex() �����s�������ƁB
' =				  �Ȃ��AlFileInfoTagIndex �� Folder �I�u�W�F�N�g GetDetailsOf()
' =				  �v���p�e�B�̗v�f�ԍ��ɑΉ�����B
' =				  ���蓖�Ă��Ă��Ȃ��擾����ʔԍ����w�肵���ꍇ�A
' =				  �擾���� False ��ԋp����B
' =			(��2) �G���[�ڍׂ͈ȉ��̎�ނ�����B
' =					  Success!			   : �擾����
' =					  File is not exist!   : �t�@�C����������Ȃ�
' =					  Get info type error! : �t�@�C���ڍ׏��^�C�g����������Ȃ�
' ==================================================================
Public Function GetFileDetailInfo( _
	ByVal sTrgtPath, _
	ByVal lFileInfoTagIndex, _
	ByRef vFileInfoValue, _
	ByRef vFileInfoTitle, _
	ByRef sErrorDetail _
)
	GetFileDetailInfo = True
	sErrorDetail = "Success!"
	
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(sTrgtPath) Then
		'Do Nothing
	Else
		GetFileDetailInfo = False
		sErrorDetail = "File is not exist!"
		Exit Function
	End If
	
	Dim sTrgtFolderPath
	Dim sTrgtFileName
	sTrgtFolderPath = Mid(sTrgtPath, 1, InStrRev(sTrgtPath, "\") - 1)
	sTrgtFileName = Mid(sTrgtPath, InStrRev(sTrgtPath, "\") + 1, Len(sTrgtPath))
	
	Dim objFolder
	Set objFolder = CreateObject("Shell.Application").Namespace(sTrgtFolderPath & "\")
	Dim objFile
	Set objFile = objFolder.ParseName(sTrgtFileName)
	
	If objFile Is Nothing Then
		GetFileDetailInfo = False
		sErrorDetail = "File is not exist!"
		Exit Function
	Else
		'Do Nothing
	End If
	
	vFileInfoValue = objFolder.GetDetailsOf(objFile, lFileInfoTagIndex)
	vFileInfoTitle = objFolder.GetDetailsOf("", lFileInfoTagIndex)
	If vFileInfoTitle = "" Then
		GetFileDetailInfo = False
		sErrorDetail = "Get info type error!"
		Exit Function
	Else
		'Do Nothing
	End If
End Function
'	Call Test_GetFileDetailInfo()
	Private Sub Test_GetFileDetailInfo()
		Dim sBuf
		Dim bRet
		Dim vFileInfoValue
		Dim vFileInfoTitle
		Dim sErrorDetail
		sBuf = ""
		Dim sTrgtPath
		sTrgtPath = "C:\codes\vbs\_lib\FileSystem.vbs"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFileDetailInfo(sTrgtPath, 1, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue & "�F" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 2, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue & "�F" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 3, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue & "�F" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 4, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue & "�F" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 52, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue & "�F" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 500, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "	" & vFileInfoTitle & "�F" & vFileInfoValue & "�F" & sErrorDetail
		sTrgtPath = "C:\test.txt"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFileDetailInfo(sTrgtPath, 1, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue & "�F" & sErrorDetail
		MsgBox sBuf
	End Sub

'GetDetailsOf() �͐ݒ�ł��Ȃ����߁A�������Ȃ�
Public Function SetFileDetailInfo( _
	ByVal sTrgtPath, _
	ByVal lGetInfoType, _
	ByVal sFileInfoValue _
)
	'Do Nothing
End Function

' ==================================================================
' = �T�v	�t�@�C���ڍ׏��̃C���f�b�N�X�擾
' = ����	vFileInfoTitle		Variant		[in]	�t�@�C���ڍ׏��^�C�g��
' = ����	lFileInfoTagIndex	Long		[out]	�擾����ʔԍ�(��1)
' = �ߒl						Boolean				�擾����
' = �o��	(��1) �擾�ł�����͂n�r�̃o�[�W�����ɂ���ĈقȂ�B
' =				  lFileInfoTagIndex �� Folder �I�u�W�F�N�g GetDetailsOf()
' =				  �v���p�e�B�̗v�f�ԍ��ɑΉ�����B
' =			�w�肵���^�C�g����������Ȃ��ꍇ�AFalse ��ԋp����B
' =			�������A���̃G���[��lTagInfoIndexMax�����������Ƃ����R��
' =			��������\�������顂��̏ꍇ�lTagInfoIndexMax ���\����
' =			�傫�����Ď��s���邱�ơ
' ==================================================================
Public Function GetFileDetailInfoIndex( _
	ByRef vFileInfoTitle, _
	ByRef lFileInfoTagIndex _
)
	Const lTagInfoIndexMax = 999
	
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim sTrgtFolderPath
	Dim sTrgtFileName
	sTrgtFolderPath = objFSO.GetDriveName(CreateObject("WScript.Shell").SpecialFolders("Desktop"))
	sTrgtFileName = ""
	
	Dim objFolder
	Dim objFile
	Set objFolder = CreateObject("Shell.Application").Namespace(sTrgtFolderPath & "\")
	Set objFile = objFolder.ParseName(sTrgtFileName)
	
	GetFileDetailInfoIndex = False
	lFileInfoTagIndex = lTagInfoIndexMax + 1
	Dim i
	For i = 0 To lTagInfoIndexMax
		Dim vGetTitle
		vGetTitle = objFolder.GetDetailsOf("", i)
		If vGetTitle = vFileInfoTitle Then
			lFileInfoTagIndex = i
			GetFileDetailInfoIndex = True
			Exit For
		Else
			'Do Nothing
		End If
	Next
End Function
'	Call Test_GetFileDetailInfoIndex()
	Private Sub Test_GetFileDetailInfoIndex()
		Dim lFileInfoTagIndex
		Dim bRet
		Dim sResult
		sResult = ""
		bRet = GetFileDetailInfoIndex( _
			"�^�C�g��", _
			lFileInfoTagIndex _
		)
		sResult = sResult & vbNewLine & bRet & " : " & lFileInfoTagIndex
		
		bRet = GetFileDetailInfoIndex( _
			"�B�e����", _
			lFileInfoTagIndex _
		)
		sResult = sResult & vbNewLine & bRet & " : " & lFileInfoTagIndex
		
		bRet = GetFileDetailInfoIndex( _
			"�Í����̏��", _
			lFileInfoTagIndex _
		)
		sResult = sResult & vbNewLine & bRet & " : " & lFileInfoTagIndex
		
		bRet = GetFileDetailInfoIndex( _
			"aaa", _
			lFileInfoTagIndex _
		)
		sResult = sResult & vbNewLine & bRet & " : " & lFileInfoTagIndex
		
		MsgBox sResult
	End Sub

'*********************************************************************
'* ���[�J���֐���`
'*********************************************************************
Private Function GetFolderNotExistPath( _
	ByVal sTrgtPath _
)
	Dim lIdx
	Dim objFSO
	Dim sCreDirPath
	Dim bIsTrgtPathExists
	lIdx = 0
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	sCreDirPath = sTrgtPath
	bIsTrgtPathExists = False
	Do While objFSO.FolderExists( sCreDirPath )
		bIsTrgtPathExists = True
		lIdx = lIdx + 1
		sCreDirPath = sTrgtPath & "_" & String( 3 - len(lIdx), "0" ) & lIdx
	Loop
	If bIsTrgtPathExists = True Then
		GetFolderNotExistPath = sCreDirPath
	Else
		GetFolderNotExistPath = ""
	End If
End Function
'	Call Test_GetFolderNotExistPath()
	Private Sub Test_GetFolderNotExistPath()
		Dim sOutStr
		sOutStr = ""
		sOutStr = sOutStr & vbNewLine & "*** test start! ***"
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath("C:\codes\vba")
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas")
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas")
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas")
		sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
		MsgBox sOutStr
	End Sub

Private Function GetFileNotExistPath( _
	ByVal sTrgtPath _
)
	Dim lIdx
	Dim objFSO
	Dim sFileParDirPath
	Dim sFileBaseName
	Dim sFileExtName
	Dim sCreFilePath
	Dim bIsTrgtPathExists
	
	lIdx = 0
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	sCreFilePath = sTrgtPath
	bIsTrgtPathExists = False
	Do While objFSO.FileExists( sCreFilePath )
		bIsTrgtPathExists = True
		lIdx = lIdx + 1
		sFileParDirPath = objFSO.GetParentFolderName( sTrgtPath )
		sFileBaseName = objFSO.GetBaseName( sTrgtPath ) & "_" & String( 3 - len(lIdx), "0" ) & lIdx
		sFileExtName = objFSO.GetExtensionName( sTrgtPath )
		If sFileExtName = "" Then
			sCreFilePath = sFileParDirPath & "\" & sFileBaseName
		Else
			sCreFilePath = sFileParDirPath & "\" & sFileBaseName & "." & sFileExtName
		End If
	Loop
	If bIsTrgtPathExists = True Then
		GetFileNotExistPath = sCreFilePath
	Else
		GetFileNotExistPath = ""
	End If
End Function
'	Call Test_GetFileNotExistPath()
	Private Sub Test_GetFileNotExistPath()
		Dim sOutStr
		sOutStr = ""
		sOutStr = sOutStr & vbNewLine & "*** test start! ***"
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba")
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas")
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas")
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas")
		sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
		MsgBox sOutStr
	End Sub

'ShowFileSelectDialog() �� ShowFilesSelectDialog() �p�̊֐�
'�_�C�A���O�̃t�B���^��ǉ�����B�w����@�͈ȉ��B
'  ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
'	   �E�g���q����������ꍇ�́A";"�ŋ�؂�
'	   �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
'	   �E�t�B���^����������ꍇ�A","�ŋ�؂�
'sFilters ���󕶎��̏ꍇ�A�t�B���^���N���A����B
Private Function SetDialogFilters( _
	ByVal sFilters, _
	ByRef fdDialog _
)
	fdDialog.Filters.Clear
	If sFilters = "" Then
		'Do Nothing
	Else
		Dim vFilter
		If InStr(sFilters, ",") > 0 Then
			Dim vFilters
			vFilters = Split(sFilters, ",")
			Dim lFilterIdx
			For lFilterIdx = 0 To UBound(vFilters)
				If InStr(vFilters(lFilterIdx), "/") > 0 Then
					vFilter = Split(vFilters(lFilterIdx), "/")
					If UBound(vFilter) = 1 Then
						fdDialog.Filters.Add vFilter(0), vFilter(1), lFilterIdx + 1
					Else
						MsgBox _
							"�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
							"""/"" �͈�����w�肵�Ă�������" & vbNewLine & _
							"  " & vFilters(lFilterIdx)
						MsgBox "�����𒆒f���܂��B"
						WScript.Quit
					End If
				Else
					MsgBox _
						"�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
						"��ʂƊg���q�� ""/"" �ŋ�؂��Ă��������B" & vbNewLine & _
						"  " & vFilters(lFilterIdx)
					MsgBox "�����𒆒f���܂��B"
					WScript.Quit
				End If
			Next
		Else
			If InStr(sFilters, "/") > 0 Then
				vFilter = Split(sFilters, "/")
				If UBound(vFilter) = 1 Then
					fdDialog.Filters.Add vFilter(0), vFilter(1), 1
				Else
					MsgBox _
						"�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
						"""/"" �͈�����w�肵�Ă�������" & vbNewLine & _
						"  " & sFilters
					MsgBox "�����𒆒f���܂��B"
					WScript.Quit
				End If
			Else
				MsgBox _
					"�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
					"��ʂƊg���q�� ""/"" �ŋ�؂��Ă��������B" & vbNewLine & _
					"  " & sFilters
				MsgBox "�����𒆒f���܂��B"
				WScript.Quit
			End If
		End If
	End If
End Function

'�e�X�g�p
Private Function FileSysem_Include( _
	ByVal sOpenFile _
	)
	Dim objFSO
	Dim objVbsFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
	
	ExecuteGlobal objVbsFile.ReadAll()
	objVbsFile.Close
	
	Set objVbsFile = Nothing
	Set objFSO = Nothing
End Function

' ******************************************************************
' *** �}�N��
' ******************************************************************
'GetDetailsOf()�̏ڍ׏��i�v�f�ԍ��A�^�C�g�����A�^���A�f�[�^�j�̈ꗗ��
'�f�X�N�g�b�v�z���ɏo�͂���
'Call GetDetailsOfGetDetailsOf()
Public Sub GetDetailsOfGetDetailsOf()
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim sTrgtDirPath
	Dim sTrgtFileName
	sTrgtDirPath = objFSO.GetDriveName(CreateObject("WScript.Shell").SpecialFolders("Desktop"))
	sTrgtFileName = ""
	
	Dim sLogFilePath
	sLogFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\file_tag_infos.txt"
	
	Dim objFolder
	Set objFolder = CreateObject("Shell.Application").Namespace(sTrgtDirPath & "\")
	Dim objFile
	Set objFile = objFolder.ParseName(sTrgtFileName)
	
	Dim objTxtFile
	Set objTxtFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(sLogFilePath, 2, True)
	objTxtFile.WriteLine "[Idx] " & Chr(9) & "[TypeName]" & Chr(9) & "[Title]"
	Dim i
	For i = 0 To 400
		objTxtFile.WriteLine _
			i & Chr(9) & _
			TypeName(objFolder.GetDetailsOf(objFile, i)) & Chr(9) & _
			objFolder.GetDetailsOf("", i)
	Next
	objTxtFile.Close
	
	Set objTxtFile = Nothing
	Set objFolder = Nothing
	Set objFile = Nothing
	
	CreateObject("WScript.Shell").Run "%comspec% /c """ & sLogFilePath & """"
End Sub

