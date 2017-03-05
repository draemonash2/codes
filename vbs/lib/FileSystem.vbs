Option Explicit

'*********************************************************************
'* グローバル関数定義
'*********************************************************************
'lFileListType）0：両方、1:ファイル、2:フォルダ、それ以外：格納しない
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
	
	'*** フォルダパス格納 ***
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
	
	'フォルダ内のサブフォルダを列挙
	'（サブフォルダがなければループ内は通らない）
	For Each objSubFolder In objFolder.SubFolders
		Call GetFileList( objSubFolder, asFileList, lFileListType)
	Next
	
	'*** ファイルパス格納 ***
	For Each objFile In objFolder.Files
		Select Case lFileListType
			Case 0:    bExecStore = True
			Case 1:    bExecStore = True
			Case 2:    bExecStore = False
			Case Else: bExecStore = False
		End Select
		If bExecStore = True Then
			'本スクリプトファイルは格納対象外
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

'Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
'asFileList は配列型ではなくバリアント型として定義する必要があることに注意！
'lFileListType）0：両方、1:ファイル、2:フォルダ、それ以外：格納しない
Public Function GetFileList2( _
	ByVal sTrgtDir, _
	ByRef asFileList, _
	ByVal lFileListType _
)
	Dim objFSO	'FileSystemObjectの格納先
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	
	'Dir コマンド実行（出力結果を一時ファイルに格納）
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
			sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '末尾に改行が付与されてしまうため、削除
			asFileList = Split( sTextAll, vbNewLine )
			objFile.Close
		Else
			WScript.Echo "ファイルが開けません: " & Err.Description
		End If
		Set objFile = Nothing	'オブジェクトの破棄
	Else
		WScript.Echo "エラー " & Err.Description
	End If	
	objFSO.DeleteFile sTmpFilePath, True
	Set objFSO = Nothing	'オブジェクトの破棄
	On Error Goto 0
End Function

'フォルダが既に存在している場合は何もしない
Public Function CreateDirectry( _
	ByVal sDirPath _
)
	Dim sParentDir
	Dim oFileSys
	
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	
	sParentDir = oFileSys.GetParentFolderName(sDirPath)
	
	'親ディレクトリが存在しない場合、再帰呼び出し
	If oFileSys.FolderExists( sParentDir ) = False Then
		Call CreateDirectry( sParentDir )
	End If
	
	'ディレクトリ作成
	If oFileSys.FolderExists( sDirPath ) = False Then
		oFileSys.CreateFolder sDirPath
	End If
	
	Set oFileSys = Nothing
End Function

'戻り値）1：ファイル、2、フォルダー、0：エラー（存在しないパス）
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
		GetFileOrFolder = 1 'ファイル
	ElseIf bFolderExists = True And bFileExists = False Then
		GetFileOrFolder = 2 'フォルダー
	Else
		GetFileOrFolder = 0 'エラー（存在しないパス）
	End If
End Function
'	Call Test_GetFileOrFolder()
	Private Sub Test_GetFileOrFolder()
		Dim objWshShell
		Dim sCurDir
		Set objWshShell = WScript.CreateObject( "WScript.Shell" )
		sCurDir = objWshShell.CurrentDirectory
		Call Test_GetFileOrFolder_Include( sCurDir & "\Array.vbs" )
		Call Test_GetFileOrFolder_Include( sCurDir & "\iTunes.vbs" )
		Call Test_GetFileOrFolder_Include( sCurDir & "\ProgressBar.vbs" )
		Call Test_GetFileOrFolder_Include( sCurDir & "\StopWatch.vbs" )
		Call Test_GetFileOrFolder_Include( sCurDir & "\String.vbs" )
		
		Dim oStpWtch
		
		Set oStpWtch = New StopWatch
		
		oStpWtch.StartT
	'	Dim asFileList()
	'	ReDim asFileList(-1)
	'	Call GetFileList( "Z:\300_Musics", asFileList, 0 )
		Dim asFileList
		Call GetFileList2( "Z:\300_Musics", asFileList, 1 )
		oStpWtch.StopT
		
		MsgBox oStpWtch.ElapsedTime
		Call OutputAllElement2LogFile(asFileList)
	End Sub
	Private Function Test_GetFileOrFolder_Include( _
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

'指定フォルダパスに含まれるフォルダが空か判定し、空フォルダなら削除する。
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
		
		'サブフォルダ精査
		Dim objSubFolder
		For Each objSubFolder In objFolder.SubFolders
			bIsFileFolderExists = True
			Exit For
		Next
		
		'サブファイル精査
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
			sRetStr = sRetStr & vbNewLine & "[Folder] empty  / delete / ↓ / " & sTrgtPath
			sRetStr = sRetStr & DeleteEmptyFolder( sTrgtParentDirPath )
		End If
	ElseIf objFSO.FileExists( sTrgtPath ) Then
		sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtPath )
		sRetStr = sRetStr & vbNewLine & "[File  ]        / stay   / ↓ / " & sTrgtPath
		sRetStr = sRetStr & DeleteEmptyFolder( sTrgtParentDirPath )
	Else
		sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtPath )
		sRetStr = sRetStr & vbNewLine & "[-     ]        / stay   / ↓ / " & sTrgtPath
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

'指定パスが存在する場合、"_XXX" を付与して返却する
'lFileDirType ) 1:file、2:folder、other:both
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

'*********************************************************************
'* ローカル関数定義
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
