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
' = 概要	ファイル/フォルダパス一覧を取得する
' = 引数	sTrgtDir		String		[in]	対象フォルダ
' = 引数	cFileList		Collections [out]	ファイル/フォルダパス一覧
' = 引数	lFileListType	Long		[in]	取得する一覧の形式
' =													0：両方
' =													1:ファイル
' =													2:フォルダ
' =													それ以外：格納しない
' = 戻値	なし
' = 覚書	・Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
' =			・Arrayコレクションに格納する
' ==================================================================
Public Function GetFileList3( _
	ByVal sTrgtDir, _
	ByRef cFileList, _
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
			Dim vFileList
			vFileList = Split( sTextAll, vbNewLine )
			Dim sFilePath
			For Each sFilePath In vFileList
				cFileList.add sFilePath
			Next
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
		sRetStr = sRetStr & vbNewLine & "[File	]		 / stay   / ↓ / " & sTrgtPath
		sRetStr = sRetStr & DeleteEmptyFolder( sTrgtParentDirPath )
	Else
		sTrgtParentDirPath = objFSO.GetParentFolderName( sTrgtPath )
		sRetStr = sRetStr & vbNewLine & "[-		]		 / stay   / ↓ / " & sTrgtPath
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
' = 引数	sTrgtPath		String		[in]	対象パス
' = 引数	sAddedPath		String		[out]	付与後のパス
' = 引数	lAddedPathType	Long		[out]	付与後のパス種別
' =												  1: ファイル
' =												  2: フォルダ
' = 戻値					Boolean				取得結果
' = 覚書	本関数では、ファイル/フォルダは作成しない。
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
	fdDialog.Title = "フォルダを選択してください（空欄の場合は親フォルダが選択されます）"
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

' ==================================================================
' = 概要	ドライブ情報取得（ドライブレター指定）
' = 引数	sDriveLetter	String		[in]	ドライブレター
' = 引数	lGetInfoType	Long		[in]	取得情報種別
' =													1) ボリュームラベル
' =													2) フォルダ
' =													3) ルートフォルダ
' =													4) 種類
' =													5) ファイルシステム
' =													6) 容量
' =													7) 空き領域
' =													8) シリアルナンバー
' = 引数	sDriveInfo		String		[out]	ドライブ情報
' = 戻値					Boolean				取得結果
' = 覚書	・ネットワークドライブも検索可能
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
				Case 1:		sDriveInfo = objDrive.VolumeName					'ボリュームラベル
				Case 2:		sDriveInfo = objDrive.Path							'フォルダ
				Case 3:		sDriveInfo = objDrive.RootFolder					'ルートフォルダ
				Case 4:		sDriveInfo = DRIVE_TYPE_TABLE(objDrive.DriveType)	'種類
				Case 5:		sDriveInfo = objDrive.FileSystem					'ファイルシステム
				Case 6:		sDriveInfo = FormatNumber(objDrive.TotalSize, 0)	'容量
				Case 7:		sDriveInfo = FormatNumber(objDrive.FreeSpace, 0)	'空き領域
				Case 8:		sDriveInfo = Hex(objDrive.SerialNumber)				'シリアルナンバー
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
																 sBuf = sBuf & vbNewLine &		  "*** C ドライブ ***"
		bRet = GetDriveInfoFromDriveLetter("C", 1, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ボリュームラベル："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 2, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  フォルダ："			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ルートフォルダ："		& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 4, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  種類："				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 5, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイルシステム："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 6, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  容量："				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 7, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  空き領域："			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 8, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  シリアルナンバー："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("C", 9, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  "						& sDriveInfo
																 sBuf = sBuf & vbNewLine &		  "*** X ドライブ ***"
		bRet = GetDriveInfoFromDriveLetter("X", 1, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ボリュームラベル："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 2, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  フォルダ："			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ルートフォルダ："		& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 4, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  種類："				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 5, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイルシステム："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 6, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  容量："				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 7, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  空き領域："			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 8, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  シリアルナンバー："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("X", 9, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  "						& sDriveInfo
																 sBuf = sBuf & vbNewLine &		  "*** Z ドライブ ***"
		bRet = GetDriveInfoFromDriveLetter("Z", 1, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ボリュームラベル："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 2, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  フォルダ："			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ルートフォルダ："		& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 4, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  種類："				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 5, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ファイルシステム："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 6, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  容量："				& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 7, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  空き領域："			& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 8, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  シリアルナンバー："	& sDriveInfo
		bRet = GetDriveInfoFromDriveLetter("Z", 9, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  "						& sDriveInfo
																 sBuf = sBuf & vbNewLine &		  "*** E ドライブ ***"
		bRet = GetDriveInfoFromDriveLetter("E", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "  ルートフォルダ："		& sDriveInfo
		MsgBox sBuf
	End Sub

' ==================================================================
' = 概要	ドライブ情報取得（ボリュームラベル指定）
' = 引数	sVolumeLabel	String		[in]	ボリュームラベル
' = 引数	lGetInfoType	Long		[in]	取得情報種別
' =													1) ボリュームラベル
' =													2) フォルダ
' =													3) ルートフォルダ
' =													4) 種類
' =													5) ファイルシステム
' =													6) 容量
' =													7) 空き領域
' =													8) シリアルナンバー
' = 引数	sDriveInfo		String		[out]	ドライブ情報
' = 戻値					Boolean				取得結果
' = 覚書	・sVolumeLabel が空文字列の場合、ボリュームラベルが設定
' =			  されていないドライブの情報が返却される
' =			・ネットワークドライブも検索可能
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
							Case 1:		sDriveInfo = objDrive.VolumeName					'ボリュームラベル
							Case 2:		sDriveInfo = objDrive.Path							'フォルダ
							Case 3:		sDriveInfo = objDrive.RootFolder					'ルートフォルダ
							Case 4:		sDriveInfo = DRIVE_TYPE_TABLE(objDrive.DriveType)	'種類
							Case 5:		sDriveInfo = objDrive.FileSystem					'ファイルシステム
							Case 6:		sDriveInfo = FormatNumber(objDrive.TotalSize, 0)	'容量
							Case 7:		sDriveInfo = FormatNumber(objDrive.FreeSpace, 0)	'空き領域
							Case 8:		sDriveInfo = Hex(objDrive.SerialNumber)				'シリアルナンバー
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
																			sBuf = sBuf & vbNewLine &		 "*** ドライブ名 SD256G ***"
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 1, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	ボリュームラベル：" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 2, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	フォルダ："			& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 3, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	ルートフォルダ："	& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 4, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	種類："				& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 5, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	ファイルシステム：" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 6, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	容量："				& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 7, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	空き領域："			& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 8, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	シリアルナンバー：" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("SD256G", 9, sDriveInfo) :		sBuf = sBuf & vbNewLine & bRet & "	"					& sDriveInfo
																			sBuf = sBuf & vbNewLine &		 "*** ドライブ名 logitechdd3t ***"
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 1, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	ボリュームラベル：" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 2, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	フォルダ："			& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 3, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	ルートフォルダ："	& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 4, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	種類："				& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 5, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイルシステム：" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 6, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	容量："				& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 7, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	空き領域："			& sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 8, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	シリアルナンバー：" & sDriveInfo
		bRet = GetDriveInfoFromVolumeLabel("logitechdd3t", 9, sDriveInfo) : sBuf = sBuf & vbNewLine & bRet & "	"					& sDriveInfo
																			sBuf = sBuf & vbNewLine &		 "*** ドライブ名 - ***"
		bRet = GetDriveInfoFromVolumeLabel("-", 3, sDriveInfo) :			sBuf = sBuf & vbNewLine & bRet & "	ルートフォルダ："	& sDriveInfo
																			sBuf = sBuf & vbNewLine &		 "*** ドライブ名 """" ***"
		bRet = GetDriveInfoFromVolumeLabel("", 3, sDriveInfo) :				sBuf = sBuf & vbNewLine & bRet & "	ルートフォルダ："	& sDriveInfo
		MsgBox sBuf
	End Sub

' ==================================================================
' = 概要	ファイル情報取得
' = 引数	sTrgtPath		String		[in]	ファイルパス
' = 引数	lGetInfoType	Long		[in]	取得情報種別 (※1)
' = 引数	vFileInfo		Variant		[out]	ファイル情報 (※1)
' = 戻値					Boolean				取得結果
' = 覚書	以下、参照。
' =		(※1) ファイル情報
' =			[引数]	[説明]					[プロパティ名]		[データ型]				[Get/Set]	[出力例]
' =			1		ファイル名				Name				vbString	文字列型	Get/Set		03 Ride Featuring Tony Matterhorn.MP3
' =			2		ファイルサイズ			Size				vbLong		長整数型	Get			4286923
' =			3		ファイル種類			Type				vbString	文字列型	Get			MPEG layer 3
' =			4		ファイル格納先ドライブ	Drive				vbString	文字列型	Get			Z:
' =			5		ファイルパス			Path				vbString	文字列型	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =			6		親フォルダ				ParentFolder		vbString	文字列型	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =			7		MS-DOS形式ファイル名	ShortName			vbString	文字列型	Get			03 Ride Featuring Tony Matterhorn.MP3
' =			8		MS-DOS形式パス			ShortPath			vbString	文字列型	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =			9		作成日時				DateCreated			vbDate		日付型		Get			2015/08/19 0:54:45
' =			10		アクセス日時			DateLastAccessed	vbDate		日付型		Get			2016/10/14 6:00:30
' =			11		更新日時				DateLastModified	vbDate		日付型		Get			2016/10/14 6:00:30
' =			12		属性					Attributes			vbLong		長整数型	(※2)		32
' =		(※2) 属性
' =			[値]				[説明]										[属性名]	[Get/Set]
' =			1  （0b00000001）	読み取り専用ファイル						ReadOnly	Get/Set
' =			2  （0b00000010）	隠しファイル								Hidden		Get/Set
' =			4  （0b00000100）	システム・ファイル							System		Get/Set
' =			8  （0b00001000）	ディスクドライブ・ボリューム・ラベル		Volume		Get
' =			16 （0b00010000）	フォルダ／ディレクトリ						Directory	Get
' =			32 （0b00100000）	前回のバックアップ以降に変更されていれば1	Archive		Get/Set
' =			64 （0b01000000）	リンク／ショートカット						Alias		Get
' =			128（0b10000000）	圧縮ファイル								Compressed	Get
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
		Case 1:		vFileInfo = objFile.Name				'ファイル名
		Case 2:		vFileInfo = objFile.Size				'ファイルサイズ
		Case 3:		vFileInfo = objFile.Type				'ファイル種類
		Case 4:		vFileInfo = objFile.Drive				'ファイル格納先ドライブ
		Case 5:		vFileInfo = objFile.Path				'ファイルパス
		Case 6:		vFileInfo = objFile.ParentFolder		'親フォルダ
		Case 7:		vFileInfo = objFile.ShortName			'MS-DOS形式ファイル名
		Case 8:		vFileInfo = objFile.ShortPath			'MS-DOS形式パス
		Case 9:		vFileInfo = objFile.DateCreated			'作成日時
		Case 10:	vFileInfo = objFile.DateLastAccessed	'アクセス日時
		Case 11:	vFileInfo = objFile.DateLastModified	'更新日時
		Case 12:	vFileInfo = objFile.Attributes			'属性
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
		bRet = GetFileInfo( sTrgtPath,	1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイル名："			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	2, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイルサイズ："		  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	3, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイル種類："			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	4, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイル格納先ドライブ：" & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	5, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイルパス："			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	6, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	親フォルダ："			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	7, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS形式ファイル名："   & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	8, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS形式パス："		  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	9, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	作成日時："				  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 10, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	アクセス日時："			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 11, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	更新日時："				  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 12, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	属性："					  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 13, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	："						  & vFileInfo
		sTrgtPath = "C:\codes\vbs\_lib\dummy.vbs"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFileInfo( sTrgtPath,	1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイル名："			  & vFileInfo
		MsgBox sBuf
	End Sub

'ファイル情報は「ファイル名」「属性」が設定可能
'しかし、以下のメソッドにて変更可能なため、実装しない
'  ファイル名： objFSO.MoveFile
'  属性： objFSO.GetFile( "C:\codes\a.txt" ).Attributes
Public Function SetFileInfo( _
   ByVal sTrgtPath, _
   ByVal lSetInfoType, _
   ByVal vFileInfo _
)
	'Do Nothing
End Function

' ==================================================================
' = 概要	フォルダ情報取得
' = 引数	sTrgtPath		String		[in]	フォルダパス
' = 引数	lGetInfoType	Long		[in]	取得情報種別 (※1)
' = 引数	vFolderInfo		Variant		[out]	フォルダ情報 (※1)
' = 戻値					Boolean				取得結果
' = 覚書	以下、参照。
' =		(※1) フォルダ情報
' =			[引数]	[説明]					[プロパティ名]		[データ型]				[Get/Set]	[出力例]
' =			1		フォルダ名				Name				vbString	文字列型	Get/Set		Sacrifice
' =			2		フォルダサイズ			Size				vbLong		長整数型	Get			80613775
' =			3		ファイル種類			Type				vbString	文字列型	Get			ファイル フォルダー
' =			4		ファイル格納先ドライブ	Drive				vbString	文字列型	Get			Z:
' =			5		フォルダパス			Path				vbString	文字列型	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =			6		ルート フォルダ			IsRootFolder		vbBoolean	ブール型	Get			False
' =			7		MS-DOS形式ファイル名	ShortName			vbString	文字列型	Get			Sacrifice
' =			8		MS-DOS形式パス			ShortPath			vbString	文字列型	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =			9		作成日時				DateCreated			vbDate		日付型		Get			2015/08/19 0:54:44
' =			10		アクセス日時			DateLastAccessed	vbDate		日付型		Get			2015/08/19 0:54:44
' =			11		更新日時				DateLastModified	vbDate		日付型		Get			2015/04/18 3:38:36
' =			12		属性					Attributes			vbLong		長整数型	(※2)		16
' =		(※2) 属性
' =			[値]				[説明]										[属性名]	[Get/Set]
' =			1  （0b00000001）	読み取り専用ファイル						ReadOnly	Get/Set
' =			2  （0b00000010）	隠しファイル								Hidden		Get/Set
' =			4  （0b00000100）	システム・ファイル							System		Get/Set
' =			8  （0b00001000）	ディスクドライブ・ボリューム・ラベル		Volume		Get
' =			16 （0b00010000）	フォルダ／ディレクトリ						Directory	Get
' =			32 （0b00100000）	前回のバックアップ以降に変更されていれば1	Archive		Get/Set
' =			64 （0b01000000）	リンク／ショートカット						Alias		Get
' =			128（0b10000000）	圧縮ファイル								Compressed	Get
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
		Case 1:		vFolderInfo = objFolder.Name				'フォルダ名
		Case 2:		vFolderInfo = objFolder.Size				'フォルダサイズ
		Case 3:		vFolderInfo = objFolder.Type				'ファイル種類
		Case 4:		vFolderInfo = objFolder.Drive				'ファイル格納先ドライブ
		Case 5:		vFolderInfo = objFolder.Path				'フォルダパス
		Case 6:		vFolderInfo = objFolder.IsRootFolder		'ルート フォルダ
		Case 7:		vFolderInfo = objFolder.ShortName			'MS-DOS形式ファイル名
		Case 8:		vFolderInfo = objFolder.ShortPath			'MS-DOS形式パス
		Case 9:		vFolderInfo = objFolder.DateCreated			'作成日時
		Case 10:	vFolderInfo = objFolder.DateLastAccessed	'アクセス日時
		Case 11:	vFolderInfo = objFolder.DateLastModified	'更新日時
		Case 12:	vFolderInfo = objFolder.Attributes			'属性
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
		bRet = GetFolderInfo( sTrgtPath, 1,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイル名："			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 2,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイルサイズ："		  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 3,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイル種類："			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 4,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイル格納先ドライブ：" & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 5,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイルパス："			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 6,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	親フォルダ："			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 7,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS形式ファイル名："   & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 8,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS形式パス："		  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 9,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	作成日時："				  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 10, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	アクセス日時："			  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 11, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	更新日時："				  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 12, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	属性："					  & vFolderInfo
		bRet = GetFolderInfo( sTrgtPath, 13, vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	："						  & vFolderInfo
		sTrgtPath = "C:\codes\vbs\libs"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFolderInfo( sTrgtPath, 1,  vFolderInfo) : sBuf = sBuf & vbNewLine & bRet & "	ファイル名："			  & vFolderInfo
		MsgBox sBuf
	End Sub

'フォルダ情報は「ファイル名」「属性」が設定可能
'しかし、以下のメソッドにて変更可能なため、実装しない
'  ファイル名： objFSO.MoveFolder
'  属性： objFSO.GetFolder( "C:\codes" ).Attributes
Public Function SetFolderInfo( _
   ByVal sTrgtPath, _
   ByVal lSetInfoType, _
   ByVal vFolderInfo _
)
	'Do Nothing
End Function

' ==================================================================
' = 概要	ファイル詳細情報取得
' = 引数	sTrgtPath			String		[in]	ファイルパス
' = 引数	lFileInfoTagIndex	Long		[in]	取得情報種別番号(※1)
' = 引数	vFileInfoValue		Variant		[out]	ファイル詳細情報
' = 引数	vFileInfoTitle		Variant		[out]	ファイル詳細情報タイトル
' = 引数	sErrorDetail		String		[out]	取得結果エラー詳細(※2)
' = 戻値						Boolean				取得結果
' = 覚書	(※1) 取得できる情報はＯＳのバージョンによって異なる。
' =				  事前に GetFileDetailInfoIndex() を実行おくこと。
' =				  なお、lFileInfoTagIndex は Folder オブジェクト GetDetailsOf()
' =				  プロパティの要素番号に対応する。
' =				  割り当てられていない取得情報種別番号を指定した場合、
' =				  取得結果 False を返却する。
' =			(※2) エラー詳細は以下の種類がある。
' =					  Success!			   : 取得成功
' =					  File is not exist!   : ファイルが見つからない
' =					  Get info type error! : ファイル詳細情報タイトルが見つからない
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
		bRet = GetFileDetailInfo(sTrgtPath, 1, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue & "：" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 2, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue & "：" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 3, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue & "：" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 4, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue & "：" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 52, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue & "：" & sErrorDetail
		bRet = GetFileDetailInfo(sTrgtPath, 500, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "	" & vFileInfoTitle & "：" & vFileInfoValue & "：" & sErrorDetail
		sTrgtPath = "C:\test.txt"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFileDetailInfo(sTrgtPath, 1, vFileInfoValue, vFileInfoTitle, sErrorDetail): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue & "：" & sErrorDetail
		MsgBox sBuf
	End Sub

'GetDetailsOf() は設定できないため、実装しない
Public Function SetFileDetailInfo( _
	ByVal sTrgtPath, _
	ByVal lGetInfoType, _
	ByVal sFileInfoValue _
)
	'Do Nothing
End Function

' ==================================================================
' = 概要	ファイル詳細情報のインデックス取得
' = 引数	vFileInfoTitle		Variant		[in]	ファイル詳細情報タイトル
' = 引数	lFileInfoTagIndex	Long		[out]	取得情報種別番号(※1)
' = 戻値						Boolean				取得結果
' = 覚書	(※1) 取得できる情報はＯＳのバージョンによって異なる。
' =				  lFileInfoTagIndex は Folder オブジェクト GetDetailsOf()
' =				  プロパティの要素番号に対応する。
' =			指定したタイトルが見つからない場合、False を返却する。
' =			ただし、このエラーはlTagInfoIndexMaxが小さいことが理由で
' =			発生する可能性がある｡その場合､lTagInfoIndexMax を十分に
' =			大きくして実行すること｡
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
			"タイトル", _
			lFileInfoTagIndex _
		)
		sResult = sResult & vbNewLine & bRet & " : " & lFileInfoTagIndex
		
		bRet = GetFileDetailInfoIndex( _
			"撮影日時", _
			lFileInfoTagIndex _
		)
		sResult = sResult & vbNewLine & bRet & " : " & lFileInfoTagIndex
		
		bRet = GetFileDetailInfoIndex( _
			"暗号化の状態", _
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

' ******************************************************************
' *** マクロ
' ******************************************************************
'GetDetailsOf()の詳細情報（要素番号、タイトル情報、型名、データ）の一覧を
'デスクトップ配下に出力する
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

