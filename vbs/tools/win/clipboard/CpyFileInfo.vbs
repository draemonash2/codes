'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

' [Path]	[Name]	[DateLastModified]	[DateCreated]	[DateLastAccessed]	[Size]	[Type]	[Attributes]

'####################################################################
'### 設定
'####################################################################
Const INCLUDE_DOUBLE_QUOTATION = False

'####################################################################
'### 本処理
'####################################################################
Const PROG_NAME = "ファイルパス/名前/更新日時コピー"

Dim bIsContinue
bIsContinue = True

Dim cFilePaths

'*** 選択ファイル取得 ***
If bIsContinue = True Then
	If EXECUTION_MODE = 0 Then 'Explorerから実行
		Set cFilePaths = CreateObject("System.Collections.ArrayList")
		Dim sArg
		For Each sArg In WScript.Arguments
			cFilePaths.add sArg
		Next
	ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
		Set cFilePaths = WScript.Col( WScript.Env("Selected") )
	Else 'デバッグ実行
		MsgBox "デバッグモードです。"
		Set cFilePaths = CreateObject("System.Collections.ArrayList")
		cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
		cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
	End If
Else
	'Do Nothing
End If

'*** ファイルパスチェック ***
If bIsContinue = True Then
	If cFilePaths.Count = 0 Then
		MsgBox "ファイルが選択されていません", vbYes, PROG_NAME
		MsgBox "処理を中断します", vbYes, PROG_NAME
		bIsContinue = False
	Else
		'Do Nothing
	End If
Else
	'Do Nothing
End If

'*** クリップボードへコピー ***
If bIsContinue = True Then
	Dim sOutString
	Dim bFirstStore
	bFirstStore = True
	Dim oFilePath
	Dim sObjName
	Dim sModDate
	Dim sObjInfo
	Dim sObjSize
	Dim sObjType
	Dim sCreateDate
	Dim sAccessDate
	Dim sAttribute
	Dim lObjType
	For Each oFilePath In cFilePaths
		lObjType = GetFileOrFolder(oFilePath)
		Select case lObjType:
			Case 1 'File
				call GetFileInfo(oFilePath, 1, sObjName)
				call GetFileInfo(oFilePath, 2, sObjSize)
				call GetFileInfo(oFilePath, 3, sObjType)
				call GetFileInfo(oFilePath, 9, sCreateDate)
				call GetFileInfo(oFilePath, 10, sAccessDate)
				call GetFileInfo(oFilePath, 11, sModDate)
				call GetFileInfo(oFilePath, 12, sAttribute)
				sObjInfo = oFilePath & _
					vbTab & sObjName & _
					vbTab & sModDate & _
					vbTab & sCreateDate & _
					vbTab & sAccessDate & _
					vbTab & sObjSize & _
					vbTab & sObjType & _
					vbTab & sAttribute
			Case 2 'Folder
				call GetFolderInfo(oFilePath, 1, sObjName)
				call GetFolderInfo(oFilePath, 2, sObjSize)
				call GetFolderInfo(oFilePath, 3, sObjType)
				call GetFolderInfo(oFilePath, 9, sCreateDate)
				call GetFolderInfo(oFilePath, 10, sAccessDate)
				call GetFolderInfo(oFilePath, 11, sModDate)
				call GetFolderInfo(oFilePath, 12, sAttribute)
				sObjInfo = oFilePath & _
					vbTab & sObjName & _
					vbTab & sModDate & _
					vbTab & sCreateDate & _
					vbTab & sAccessDate & _
					vbTab & sObjSize & _
					vbTab & sObjType & _
					vbTab & sAttribute
			Case Else 'Not Exist
				'Do Nohting
		End Select
		If bFirstStore = True Then
			sOutString = sObjInfo
			bFirstStore = False
		Else
			sOutString = sOutString & vbNewLine & sObjInfo
		End If
	Next
	CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
	'Do Nothing
End If

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
