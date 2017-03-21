Option Explicit

'*********************************************************************
'* グローバル関数定義
'*********************************************************************
' ==================================================================
' = 概要	ファイル/フォルダパス一覧を取得する
' = 引数	sTrgtDir		String		[in]	対象フォルダ
' = 引数	asFileList		String()	[out]	ファイル/フォルダパス一覧
' = 引数	lFileListType	Long		[in]	取得する一覧の形式
' =													0：両方
' =													1:ファイル
' =													2:フォルダ
' =													それ以外：格納しない
' = 戻値	なし
' = 覚書	なし
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
' = 概要	ファイル/フォルダパス一覧を取得する
' = 引数	sTrgtDir		String		[in]	対象フォルダ
' = 引数	asFileList		Variant		[out]	ファイル/フォルダパス一覧
' = 引数	lFileListType	Long		[in]	取得する一覧の形式
' =													0：両方
' =													1:ファイル
' =													2:フォルダ
' =													それ以外：格納しない
' = 戻値	なし
' = 覚書	・Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
' =			・asFileList は配列型ではなくバリアント型として定義する
' =			  必要があることに注意！
' ==================================================================
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
' = 概要	フォルダを作成する
' = 引数	sDirPath		String		[in]	作成対象フォルダ
' = 戻値	なし
' = 覚書	・作成対象フォルダの親ディレクトリが存在しない場合、
' =			  再帰的に親フォルダを作成する
' =			・フォルダが既に存在している場合は何もしない
' ==================================================================
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

' ==================================================================
' = 概要	ファイルかフォルダかを判定する
' = 引数	sChkTrgtPath	String		[in]	チェック対象フォルダ
' = 戻値					Long				判定結果
' =													1) ファイル
' =													2) フォルダー
' =													0) エラー（存在しないパス）
' = 覚書	FileSystemObject を使っているので、ファイル/フォルダの
' =			存在確認にも使用可能。
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
		GetFileOrFolder = 1 'ファイル
	ElseIf bFolderExists = True And bFileExists = False Then
		GetFileOrFolder = 2 'フォルダー
	Else
		GetFileOrFolder = 0 'エラー（存在しないパス）
	End If
End Function

' ==================================================================
' = 概要	指定フォルダパスに含まれるフォルダが空か判定し、
' =			空フォルダなら削除する。
' = 引数	sTrgtPath	String		[in]	チェック対象フォルダ
' = 戻値				String				削除結果ログ
' = 覚書	なし
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

' ==================================================================
' = 概要	指定パスが存在する場合、"_XXX" を付与して返却する
' = 引数	sTrgtPath		String		[in]	対象フォルダ
' = 引数	lFileDirType	Long		[in]	ファイル/フォルダ種別
' =													1:ファイル
' =													2:フォルダ
' = 戻値					String				フォルダパス
' = 覚書	本関数では、ファイル/フォルダは作成しない。
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
' = 概要	フォルダ選択ダイアログを表示する
' = 引数	sInitPath	String	[in]  デフォルトフォルダパス
' = 戻値				String		  フォルダ選択結果
' = 覚書	・存在しないフォルダパスを選択した場合、空文字列を返却する
' =			・キャンセルを押下した場合、空文字列を返却する
' ==================================================================
Private Function ShowFolderSelectDialog( _
	ByVal sInitPath _
)
	Const msoFileDialogFolderPicker = 4
	Const xlMinimized = -4140
	
	Dim objExcel
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False '非表示にしても閉じる際にちらっと表示されちゃう。
	objExcel.WindowState = xlMinimized '上記理由から最小化もしとく。
	
	Dim fdDialog
	Set fdDialog = objExcel.FileDialog(msoFileDialogFolderPicker)
	fdDialog.Title = "フォルダを選択してください"
	If sInitPath = "" Then
		'Do Nothing
	Else
		If Right(sInitPath, 1) = "\" Then
			fdDialog.InitialFileName = sInitPath
		Else
			fdDialog.InitialFileName = sInitPath & "\"
		End If
	End If
	
	'ダイアログ表示
	Dim lResult
	lResult = fdDialog.Show()
	If lResult <> -1 Then 'キャンセル押下
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
' = 概要	ファイル（単一）選択ダイアログを表示する
' = 引数	sInitPath	String	[in]  デフォルトファイルパス
' = 引数	sFilters	String	[in]  選択時のフィルタ(※)
' = 戻値				String		  ファイル選択結果
' = 覚書	(※)ダイアログのフィルタ指定方法は以下。
' =				 ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =					   ・拡張子が複数ある場合は、";"で区切る
' =					   ・ファイル種別と拡張子は"/"で区切る
' =					   ・フィルタが複数ある場合、","で区切る
' =			sFilters が省略もしくは空文字の場合、フィルタをクリアする。
' ==================================================================
Private Function ShowFileSelectDialog( _
	ByVal sInitPath, _
	ByVal sFilters _
)
	Const msoFileDialogFilePicker = 3
	Const xlMinimized = -4140
	
	Dim objExcel
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False '非表示にしても閉じる際にちらっと表示されちゃう。
	objExcel.WindowState = xlMinimized '上記理由から最小化もしとく。
	
	Dim fdDialog
	Set fdDialog = objExcel.FileDialog(msoFileDialogFilePicker)
	fdDialog.Title = "ファイルを選択してください"
	fdDialog.AllowMultiSelect = False
	If sInitPath = "" Then
		'Do Nothing
	Else
		fdDialog.InitialFileName = sInitPath
	End If
	Call SetDialogFilters(sFilters, fdDialog) 'フィルタ追加
	
	'ダイアログ表示
	Dim lResult
	lResult = fdDialog.Show()
	If lResult <> -1 Then 'キャンセル押下
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
		'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png"
		'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv"
		'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
		sFilters = ""
		
		MsgBox ShowFileSelectDialog( sInitPath, sFilters )
	End Sub

' ==================================================================
' = 概要	ファイル（複数）選択ダイアログを表示する
' = 引数	asSelectedFiles String()	[out] 選択されたファイルパス一覧
' = 引数	sInitPath		String		[in]  デフォルトファイルパス
' = 引数	sFilters		String		[in]  選択時のフィルタ(※)
' = 戻値	なし
' = 覚書	(※)ダイアログのフィルタ指定方法は以下。
' =				 ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =					   ・拡張子が複数ある場合は、";"で区切る
' =					   ・ファイル種別と拡張子は"/"で区切る
' =					   ・フィルタが複数ある場合、","で区切る
' =			sFilters が省略もしくは空文字の場合、フィルタをクリアする。
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
	objExcel.Visible = False '非表示にしても閉じる際にちらっと表示されちゃう。
	objExcel.WindowState = xlMinimized '上記理由から最小化もしとく。
	
	Dim fdDialog
	Set fdDialog = objExcel.FileDialog(msoFileDialogFilePicker)
	fdDialog.Title = "ファイルを選択してください（複数可）"
	fdDialog.AllowMultiSelect = True
	If sInitPath = "" Then
		'Do Nothing
	Else
		fdDialog.InitialFileName = sInitPath
	End If
	Call SetDialogFilters(sFilters, fdDialog) 'フィルタ追加
	
	'ダイアログ表示
	Dim lResult
	lResult = fdDialog.Show()
	If lResult <> -1 Then 'キャンセル押下
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
		'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png"
		'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv"
		'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
		sFilters = "全てのファイル/*.*,画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
		
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

'ShowFileSelectDialog() と ShowFilesSelectDialog() 用の関数
'ダイアログのフィルタを追加する。指定方法は以下。
'  ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
'	   ・拡張子が複数ある場合は、";"で区切る
'	   ・ファイル種別と拡張子は"/"で区切る
'	   ・フィルタが複数ある場合、","で区切る
'sFilters が空文字の場合、フィルタをクリアする。
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
							"ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
							"""/"" は一つだけ指定してください" & vbNewLine & _
							"  " & vFilters(lFilterIdx)
						MsgBox "処理を中断します。"
						WScript.Quit
					End If
				Else
					MsgBox _
						"ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
						"種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
						"  " & vFilters(lFilterIdx)
					MsgBox "処理を中断します。"
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
						"ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
						"""/"" は一つだけ指定してください" & vbNewLine & _
						"  " & sFilters
					MsgBox "処理を中断します。"
					WScript.Quit
				End If
			Else
				MsgBox _
					"ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
					"種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
					"  " & sFilters
				MsgBox "処理を中断します。"
				WScript.Quit
			End If
		End If
	End If
End Function

'テスト用
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
