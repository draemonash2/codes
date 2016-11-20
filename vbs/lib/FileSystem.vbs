Option Explicit

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

'指定フォルダパスに含まれるフォルダが空か判定し、空フォルダなら削除する。
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
		
		'サブフォルダ精査
		Dim objSubFolder
		For Each objSubFolder In objFolder.SubFolders
			bIsFileFolderExists = True
		Next
		
		'サブファイル精査
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
