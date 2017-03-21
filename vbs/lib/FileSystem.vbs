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
		sRetStr = sRetStr & vbNewLine & "[File  ]        / stay   / �� / " & sTrgtPath
		sRetStr = sRetStr & DeleteEmptyFolder( sTrgtParentDirPath )
	Else
		sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtPath )
		sRetStr = sRetStr & vbNewLine & "[-     ]        / stay   / �� / " & sTrgtPath
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
' = ����	sTrgtPath		String		[in]	�Ώۃt�H���_
' = ����	lFileDirType	Long		[in]	�t�@�C��/�t�H���_���
' =													1:�t�@�C��
' =													2:�t�H���_
' = �ߒl					String				�t�H���_�p�X
' = �o��	�{�֐��ł́A�t�@�C��/�t�H���_�͍쐬���Ȃ��B
' ==================================================================
Public Function GetNotExistPath( _
	ByVal sTrgtPath, _
	ByVal lFileDirType _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If lFileDirType = 1 Then
		GetNotExistPath = GetFileNotExistPath( sTrgtPath )
	ElseIf lFileDirType = 2 Then
		GetNotExistPath = GetFolderNotExistPath( sTrgtPath )
	Else
		GetNotExistPath = ""
	End If
End Function
'	Call Test_GetNotExistPath()
	Private Sub Test_GetNotExistPath()
		Dim sOutStr
		sOutStr = ""
		sOutStr = sOutStr & vbNewLine & "*** test start! ***"
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\a",		0 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\a",		1 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\a",		2 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\b.txt",	0 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\b.txt",	1 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\b.txt",	2 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\c.txt",	0 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\c.txt",	1 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\c.txt",	2 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\d",		0 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\d",		1 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\d",		2 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\e",		0 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\e",		1 )
		sOutStr = sOutStr & vbNewLine & GetNotExistPath( "C:\codes\vbs\test\e",		2 )
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
	fdDialog.Title = "�t�H���_��I�����Ă�������"
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
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\a" )
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\b.txt" )
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\c.txt" )
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\d" )
		sOutStr = sOutStr & vbNewLine & GetFolderNotExistPath( "C:\codes\vbs\test\e" )
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
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\a" )
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\b.txt" )
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\c.txt" )
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\d" )
		sOutStr = sOutStr & vbNewLine & GetFileNotExistPath( "C:\codes\vbs\test\e" )
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
