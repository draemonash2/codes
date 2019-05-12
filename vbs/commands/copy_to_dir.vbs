Option Explicit

'==============================================================================
'【概要】
'	ファイル/フォルダをコピーする。
'	移動先のフォルダが存在しない場合、フォルダを作成してからコピーする。
'
'【使用方法】
'	copy_to_dir.vbs <source_path> <destination_path>
'
'【使用例】
'	1) copy_to_dir.vbs c:\codes\vbs\test.txt c:\test\test.txt
'	2) copy_to_dir.vbs c:\codes\vbs c:\test\vbs
'		c:\codes\vbs
'			└ a.txt
'			└ b
'				└ c.txt
'		↓
'		c:\test\vbs
'			└ a.txt
'			└ b
'				└ c.txt
'
'【覚え書き】
'	なし
'
'【改訂履歴】
'	1.0.0	2019/05/12	新規作成
'==============================================================================

'==============================================================================
' 設定
'==============================================================================

'==============================================================================
' 本処理
'==============================================================================
'引数チェック
If WScript.Arguments.Count = 2 Then
	'Do Nothing
Else
	Wscript.quit
End If

dim sSrcPath
dim sDstPath
sSrcPath = Replace(WScript.Arguments(0), "/", "\")
sDstPath = Replace(WScript.Arguments(1), "/", "\")

Dim lSrcPathType
lSrcPathType = GetFileOrFolder(sSrcPath)

dim sDstParDir
sDstParDir = GetDirPath( sDstPath )

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If lSrcPathType = 1 Then 'ファイル
	call CreateDirectry( sDstParDir )
	objFSO.CopyFile sSrcPath, sDstPath
ElseIf lSrcPathType = 2 Then 'フォルダ
	call CreateDirectry( sDstParDir )
	objFSO.CopyFolder sSrcPath, sDstPath
Else '未存在
'	WScript.Echo "ファイルが存在しません"
End If

Set objFSO = Nothing

'==============================================================================
' ライブラリ
'==============================================================================
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
' = 概要	指定されたファイルパスからフォルダパスを抽出する
' = 引数	sFilePath	String	[in]  ファイルパス
' = 戻値				String		  フォルダパス
' = 覚書	ローカルファイルパス（例：c:\test）や URL （例：https://test）
' =			が指定可能
' ==================================================================
Public Function GetDirPath( _
	ByVal sFilePath _
)
	If InStr( sFilePath, "\" ) Then
		GetDirPath = RemoveTailWord( sFilePath, "\" )
	ElseIf InStr( sFilePath, "/" ) Then
		GetDirPath = RemoveTailWord( sFilePath, "/" )
	Else
		GetDirPath = sFilePath
	End If
End Function
	'Call Test_GetDirPath()
	Private Sub Test_GetDirPath()
		Dim Result
		Result = "[Result]"
		Result = Result & vbNewLine & GetDirPath( "C:\test\a.txt" )    ' C:\test
		Result = Result & vbNewLine & GetDirPath( "http://test/a" )    ' http://test
		Result = Result & vbNewLine & GetDirPath( "C:_test_a.txt" )    ' C:_test_a.txt
		MsgBox Result
	End Sub

' ==================================================================
' = 概要	末尾区切り文字以降の文字列を除去する。
' = 引数	sStr		String	[in]  分割する文字列
' = 引数	sDlmtr		String	[in]  区切り文字
' = 戻値				String		  除去文字列
' = 覚書	なし
' ==================================================================
Public Function RemoveTailWord( _
	ByVal sStr, _
	ByVal sDlmtr _
)
	Dim sTailWord
	Dim lRemoveLen
	
	If sStr = "" Then
		RemoveTailWord = ""
	Else
		If sDlmtr = "" Then
			RemoveTailWord = sStr
		Else
			If InStr(sStr, sDlmtr) = 0 Then
				RemoveTailWord = sStr
			Else
				sTailWord = ExtractTailWord(sStr, sDlmtr)
				lRemoveLen = Len(sDlmtr) + Len(sTailWord)
				RemoveTailWord = Left(sStr, Len(sStr) - lRemoveLen)
			End If
		End If
	End If
End Function
	'Call Test_RemoveTailWord()
	Private Sub Test_RemoveTailWord()
		Dim Result
		Result = "[Result]"
		Result = Result & vbNewLine & "*** test start! ***"
		Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "\" )	' C:\test
		Result = Result & vbNewLine & RemoveTailWord( "C:\test\a", "\" )		' C:\test
		Result = Result & vbNewLine & RemoveTailWord( "C:\test\", "\" )			' C:\test
		Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\" )			' C:
		Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\\" )			' C:\test
		Result = Result & vbNewLine & RemoveTailWord( "", "\" )					' 
		Result = Result & vbNewLine & RemoveTailWord( "a.txt", "\" )			' a.txt（ファイル名かどうかは判断しない）
		Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "" )		' C:\test\a.txt
		Result = Result & vbNewLine & "*** test finished! ***"
		MsgBox Result
	End Sub

' ==================================================================
' = 概要	末尾区切り文字以降の文字列を返却する。
' = 引数	sStr		String	[in]  分割する文字列
' = 引数	sDlmtr		String	[in]  区切り文字
' = 戻値				String		  抽出文字列
' = 覚書	なし
' ==================================================================
Public Function ExtractTailWord( _
	ByVal sStr, _
	ByVal sDlmtr _
)
	Dim asSplitWord
	
	If Len(sStr) = 0 Then
		ExtractTailWord = ""
	Else
		ExtractTailWord = ""
		asSplitWord = Split(sStr, sDlmtr)
		ExtractTailWord = asSplitWord(UBound(asSplitWord))
	End If
End Function
	'Call Test_ExtractTailWord()
	Private Sub Test_ExtractTailWord()
		Dim Result
		Result = "[Result]"
		Result = Result & vbNewLine & "*** test start! ***"
		Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "\" )	' a.txt
		Result = Result & vbNewLine & ExtractTailWord( "C:\test\a", "\" )		' a
		Result = Result & vbNewLine & ExtractTailWord( "C:\test\", "\" )		' 
		Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\" )			' test
		Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\\" )		' C:\test
		Result = Result & vbNewLine & ExtractTailWord( "a.txt", "\" )			' a.txt
		Result = Result & vbNewLine & ExtractTailWord( "", "\" )				' 
		Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "" )	' C:\test\a.txt
		Result = Result & vbNewLine & "*** test finished! ***"
		MsgBox Result
	End Sub
