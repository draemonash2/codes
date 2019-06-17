'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

' [Path]	[Name]	[Size]	[Type]	[DateLastModified]	[DateCreated]	[DateLastAccessed]	[Attributes]

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
	Dim sFileName
	Dim sModDate
	Dim sFileInfo
	Dim sFileSize
	Dim sFileType
	Dim sCreateDate
	Dim sAccessDate
	Dim sAttribute
	For Each oFilePath In cFilePaths
		call GetFileInfo(oFilePath, 1, sFileName)
		call GetFileInfo(oFilePath, 2, sFileSize)
		call GetFileInfo(oFilePath, 3, sFileType)
		call GetFileInfo(oFilePath, 9, sCreateDate)
		call GetFileInfo(oFilePath, 10, sAccessDate)
		call GetFileInfo(oFilePath, 11, sModDate)
		call GetFileInfo(oFilePath, 12, sAttribute)
		sFileInfo = oFilePath & _
			vbTab & sFileName & _
			vbTab & sModDate & _
			vbTab & sCreateDate & _
			vbTab & sAccessDate & _
			vbTab & sFileSize & _
			vbTab & sFileType & _
			vbTab & sAttribute
		If bFirstStore = True Then
			sOutString = sFileInfo
			bFirstStore = False
		Else
			sOutString = sOutString & vbNewLine & sFileInfo
		End If
	Next
	CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
	'Do Nothing
End If

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

