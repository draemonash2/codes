Option Explicit

'ファイルを指定すると現在時刻を付与したバックアップファイルを作成する。
'同じファイル名のものが存在していたら、アルファベットを付与したバックアップファイルを作成する。
'	ex. 211201a, 211202b, …
'指定数分バックアップがたまったら、古いものから削除する。
'バックアップ対象はファイルのみ。
'第二引数にバックアップファイル数を指定できる。
'★要修正★バックアップ最大数分、同じ日付のバックアップファイルで満たされると、新しいファイルが更新されていくため要注意。

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )		'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()

'===============================================================================
'= 設定値
'===============================================================================
Const lMAX_BAK_FILE_NUM_DEFAULT = 50
Const sBAK_DIR_NAME = "_bak"
Const sBAK_FILE_SUFFIX = "#b#"

'===============================================================================
'= 本処理
'===============================================================================
Const sSCRIPT_NAME = "ファイルバックアップ"
Dim sTrgtFilePath
Dim lBakFileNumMax
If WScript.Arguments.Count >= 2 Then
	sTrgtFilePath = WScript.Arguments(0)
	lBakFileNumMax = CLng(WScript.Arguments(1))
ElseIf WScript.Arguments.Count = 1 Then
	sTrgtFilePath = WScript.Arguments(0)
	lBakFileNumMax = lMAX_BAK_FILE_NUM_DEFAULT
Else
	WScript.Echo "引数を指定してください。プログラムを中断します。"
	WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'対象ファイル情報取得
Dim sTrgtFileParDirPath
Dim sTrgtFileBaseName
Dim sTrgtFileExt
Dim sDateSuffix
sTrgtFileParDirPath = objFSO.GetParentFolderName( sTrgtFilePath )
sTrgtFileBaseName = objFSO.GetBaseName( sTrgtFilePath )
sTrgtFileExt = objFSO.GetExtensionName( sTrgtFilePath )
sDateSuffix = ConvDate2String(Now(),2)

'バックアップファイル情報作成
Dim sBakDirPath
Dim sBakFilePathBase
Dim sBakFilePath
sBakDirPath = sTrgtFileParDirPath & "\" & sBAK_DIR_NAME
sBakFilePathBase = sBakDirPath & "\" & sTrgtFileBaseName & "_" & sBAK_FILE_SUFFIX
sBakFilePath = sBakFilePathBase & sDateSuffix & "." & sTrgtFileExt

'バックアップフォルダ作成
Call CreateDirectry( sBakDirPath )

'*** ファイルバックアップ ***
'未存在ファイルパス判定
Dim lAlphaIdx
lAlphaIdx = 97 'asciiコードのa
Do While objFSO.FileExists( sBakFilePath )
	sBakFilePath = sBakFilePathBase & sDateSuffix & Chr(lAlphaIdx) & "." & sTrgtFileExt
	lAlphaIdx = lAlphaIdx + 1
Loop
'ファイルバックアップ
objFSO.CopyFile sTrgtFilePath, sBakFilePath

'*** 古いファイル削除 ***
'ファイルリスト取得
Dim cFileList
Set cFileList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct( sBakDirPath, cFileList, 1, "*")

'バックアップファイル数取得
Dim lBakFileNum
Dim sDelFilePath
lBakFileNum = 0
sDelFilePath = ""
Dim sFilePath
For Each sFilePath in cFileList
	If ( (InStr(sFilePath, sBakFilePathBase) > 0) And _
	     (objFSO.GetExtensionName(sFilePath) = sTrgtFileExt) ) Then
		 If lBakFileNum = 0 Then
			sDelFilePath = sFilePath
		 End If
		 lBakFileNum = lBakFileNum + 1
	End If
Next

'バックアップファイル削除
If lBakFileNum > lBakFileNumMax Then
	objFSO.DeleteFile sDelFilePath, True
End If

'MsgBox "バックアップ完了！", vbOKOnly, sSCRIPT_NAME

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
