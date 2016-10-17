Option Explicit

'==========================================================
'= 設定値
'==========================================================
Const TRGT_DIR = "Z:\300_Musics"
Const UPDATE_MOD_DATE = False

'==========================================================
'= 本処理
'==========================================================
Dim objWshShell
Dim sCurDir
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
sCurDir = objWshShell.CurrentDirectory
Call Include( sCurDir & "\lib\String.vbs" )
Call Include( sCurDir & "\lib\StopWatch.vbs" )
Call Include( sCurDir & "\lib\ProgressBar.vbs" )
Call Include( sCurDir & "\lib\FileSystem.vbs" )
Call Include( sCurDir & "\lib\iTunes.vbs" )
Call Include( sCurDir & "\lib\Array.vbs" )

Dim sLogFilePath
sLogFilePath = sCurDir & "\" & RemoveTailWord( WScript.ScriptName, "." ) & ".log"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objLogFile
Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )

objLogFile.WriteLine "[更新対象フォルダ] " & TRGT_DIR
objLogFile.WriteLine "[更新日時 変更有無] " & UPDATE_MOD_DATE

Dim oStpWtch
Set oStpWtch = New StopWatch
Call oStpWtch.StartT

Dim oPrgBar
Set oPrgBar = New ProgressBar

If 1 Then ' ★Debug★

' ******************************************
' * 日付入力                               *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** 日付入力処理 *** "
oPrgBar.SetMsg( _
	"⇒・日付入力処理" & vbNewLine & _
	"　・更新対象ファイル特定処理" & vbNewLine & _
	"　・タグ更新処理" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '進捗更新

On Error Resume Next

oPrgBar.SetProg( 10 ) '進捗更新

Dim sNow
sNow = Now()
sNow = Left( sNow, Len( sNow ) - 2 ) & "00" '秒を00にする

Dim sCmpBaseTime
sCmpBaseTime = InputBox( _
					"更新対象とするファイルを特定します。" & vbNewLine & _
					"更新対象とする時刻を入力してください。" & vbNewLine & _
					"" & vbNewLine & _
					"  [入力規則] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
					"" & vbNewLine & _
					"※ 日付のみを指定したい場合、「YYYY/MM/DD 0:0:0」としてください。" _
					, "入力" _
					, sNow _
				)

objLogFile.WriteLine "入力値 : " & sCmpBaseTime

Dim sTimeValue
Dim sDateValue
sTimeValue = TimeValue(sCmpBaseTime)
sDateValue = DateValue(sCmpBaseTime)

oPrgBar.SetProg( 50 ) '進捗更新

'日付チェック
If Err.Number <> 0 Then
	MsgBox "日付の形式が不正です！" & vbNewLine & _
	       "  [入力規則] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
	       "  [入力値] " & sCmpBaseTime
	MsgBox Err.Description
	MsgBox "プログラムを中断します！"
	Err.Clear
	Call Finish
	WScript.Quit
Else
	'Do Nothing
End If
If DateDiff("s", sCmpBaseTime, Now() ) < 0  Then
	MsgBox "未来の日時が指定されました！" & vbNewLine & _
	       "  [入力値] " & sCmpBaseTime
	MsgBox "プログラムを中断します！"
	Call Finish
	WScript.Quit
Else
	'Do Nothing
End If
On Error Goto 0 '「On Error Resume Next」を解除


oPrgBar.SetProg( 100 ) '進捗更新

objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime & " [s]"

' ******************************************
' * 更新対象ファイルリスト取得             *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** 更新対象ファイル特定 *** "
oPrgBar.SetMsg( _
	"　・日付入力処理" & vbNewLine & _
	"⇒・更新対象ファイル特定処理" & vbNewLine & _
	"　・タグ更新処理" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '進捗更新

On Error Resume Next

'*** Dir コマンド実行 ***
Dim sTmpFilePath
Dim sExecCmd
sTmpFilePath = objWshShell.CurrentDirectory & "\" & replace( WScript.ScriptName, ".vbs", "_TrgtFileList.tmp" )
sExecCmd = "Dir """ & TRGT_DIR & """ /s /a:a-d > """ & sTmpFilePath & """"
With CreateObject("Wscript.Shell")	
	.Run "cmd /c" & sExecCmd, 7, True
End With
Dim objFile
Dim sTextAll
Dim asTxtArray
If Err.Number = 0 Then
	Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
	If Err.Number = 0 Then
		sTextAll = objFile.ReadAll
		sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '末尾に改行が付与されてしまうため、削除
		asTxtArray = Split( sTextAll, vbNewLine )
		objFile.Close
	Else
		WScript.Echo "ファイルが開けません: " & Err.Description
	End If
	Set objFile = Nothing	'オブジェクトの破棄
Else
	WScript.Echo "エラー " & Err.Description
End If
On Error Goto 0

oPrgBar.SetProg( 20 ) '進捗更新

'*** Dir コマンド結果取得 ＆ 更新対象ファイルリスト作成 ***
Dim lIdx
Dim sDirPath
Dim sTxtLine
Dim sModDate
Dim sFileName
Dim sFilePath
Dim sFileSize
Dim vSplitData
Dim asTrgtFileList()
ReDim asTrgtFileList(-1)
sDirPath = ""
For lIdx = 0 to UBound( asTxtArray )
	oPrgBar.SetProg( oPrgBar.ConvProgRange( 0, UBound( asTxtArray ), lIdx ) ) '進捗更新
	sTxtLine = asTxtArray( lIdx )
	If InStr( sTxtLine, " のディレクトリ" ) > 0 Then
		sDirPath = Mid( sTxtLine, 2, Len( sTxtLine ) - Len( " のディレクトリ" ) - 1 )
	ElseIf InStr( sTxtLine, "ボリューム ラベル" ) > 0 Or _
		   InStr( sTxtLine, "ボリューム シリアル番号は " ) > 0 Or _
		   ( ( InStr( sTxtLine, " 個のファイル" ) > 0 ) And ( InStr( sTxtLine, " バイト" ) > 0 ) ) Or _
		   ( ( InStr( sTxtLine, " 個のディレクトリ" ) > 0 ) And ( InStr( sTxtLine, " バイト" ) > 0 ) ) Or _
		   InStr( sTxtLine, "     ファイルの総数:" ) > 0 Or _
		   sTxtLine = "" Then
		'Do Nothing
	Else
		Do
			sTxtLine = Replace( sTxtLine, "  ", " " )
		Loop While InStr( sTxtLine, "  " ) > 0
		vSplitData = Split( sTxtLine, " " )
		If UBound( vSplitData ) < 3 Then
			MsgBox "エラー！" & vbNewLine & sTxtLine
			WScript.Quit
		End If
		sModDate = vSplitData(0) & " " & vSplitData(1) & ":00"
		sFileSize = vSplitData(2)
		Dim i
		For i = 3 to UBound( vSplitData )
			If i = 3 Then
				sFileName = vSplitData(i)
			Else
				sFileName = sFileName & " " & vSplitData(i)
			End If
		Next
		sFilePath = sDirPath & "\" & sFileName
		
		If DateDiff("s", sCmpBaseTime, sModDate ) >= 0  Then
			ReDim Preserve asTrgtFileList( UBound( asTrgtFileList ) + 1 )
			asTrgtFileList( UBound( asTrgtFileList ) ) = sFilePath
		Else
			'Do Nothing
		End If
		
		If Err.Number = 0 Then
			'Do Nothing
		Else
			MsgBox "sFilePath  : " & sFilePath & vbNewLine & _
				   "sModDate   : " & sModDate  & vbNewLine & _
				   "sFileSize  : " & sFileSize & vbNewLine & _
				   "sFileName  : " & sFileName & vbNewLine & _
				   "ArrayData  : " & asTrgtFileList( UBound( asTrgtFileList ) )
			MsgBox Err.Description
			Err.Clear
		End If
	End If
Next

objFSO.DeleteFile sTmpFilePath, True
Set objFSO = Nothing	'オブジェクトの破棄


objLogFile.WriteLine "ファイル数：" & UBound(asTrgtFileList) + 1
objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime & " [s]"

Else ' ★Debug★
	ReDim asTrgtFileList(0)
	asTrgtFileList(0) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Bow Down.mp3"
'	asTrgtFileList(1) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concentrate.mp3"
'	asTrgtFileList(2) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concrete Schoolyard.mp3"
'	asTrgtFileList(3) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Control Myself.mp3"
End If ' ★Debug★

' ******************************************
' * タグ更新                               *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** タグ更新処理 *** "
oPrgBar.SetMsg( _
	"　・日付入力処理" & vbNewLine & _
	"　・更新対象ファイル特定処理" & vbNewLine & _
	"⇒・タグ更新処理" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '進捗更新

objLogFile.WriteLine "[FilePath]" & Chr(9) & "[TrackName}" & Chr(9) & "[HitNum]"

Dim lTrgtFileListIdx
Dim lTrgtFileListNum
lTrgtFileListNum = UBound( asTrgtFileList )
For lTrgtFileListIdx = 0 To lTrgtFileListNum
	'進捗更新
	oPrgBar.SetProg( _
		oPrgBar.ConvProgRange( _
			0, _
			lTrgtFileListNum, _
			lTrgtFileListIdx _
		) _
	)
	
	Dim sTrgtFilePath
	sTrgtFilePath = asTrgtFileList( lTrgtFileListIdx )
	
	'トラック名取得
	Dim sTrgtDirPath
	Dim sTrgtFileName
	sTrgtDirPath = RemoveTailWord( sTrgtFilePath, "\" )
	sTrgtFileName = ExtractTailWord( sTrgtFilePath, "\" )
	
	Dim oTrgtDirFiles
	Dim oTrgtFile
	Dim sTrgtTrackName
	Dim sTrgtModDate
	Set oTrgtDirFiles = CreateObject("Shell.Application").Namespace( sTrgtDirPath )
	Set oTrgtFile = oTrgtDirFiles.ParseName( sTrgtFileName )
	sTrgtTrackName = oTrgtDirFiles.GetDetailsOf( oTrgtFile, 21 )
	sTrgtModDate = oTrgtFile.ModifyDate
	
	Dim objPlayList
	Dim objSearchResult
	Set objPlayList = WScript.CreateObject("iTunes.Application").Sources.Item(1).Playlists.ItemByName("ミュージック")
	Set objSearchResult = objPlayList.Search( sTrgtTrackName, 5 )
	
	objLogFile.WriteLine sTrgtFilePath & Chr(9) & sTrgtTrackName & Chr(9) & objSearchResult.Count
	
	Dim lHitIdx
	For lHitIdx = 1 to objSearchResult.Count
		With objSearchResult.Item(lHitIdx)
			If .Location = sTrgtFilePath Then
				.Composer = "1"
				.Composer = ""
				Exit For
			Else
				'Do Nothing
			End If
		End With
	Next
	Set objSearchResult = Nothing
	Set objPlayList = Nothing
	
	If UPDATE_MOD_DATE = True Then
		'Do Nothing
	Else
		oTrgtFile.ModifyDate = CDate( sTrgtModDate )
	End If
	
	Set oTrgtFile = Nothing
	Set oTrgtDirFiles = Nothing
Next

objLogFile.WriteLine "ファイル数：" & UBound(asTrgtFileList) + 1
objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime & " [s]"

' ******************************************
' * 終了処理                               *
' ******************************************
Call Finish
MsgBox "プログラムが正常に終了しました。"

'==========================================================
'= 関数定義
'==========================================================
' 外部プログラム インクルード関数
Function Include( _
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

Function Finish()
	Call oStpWtch.StopT
	Call oPrgBar.Quit
	objLogFile.WriteLine ""
	objLogFile.WriteLine "開始時刻               : " & oStpWtch.StartPoint
	objLogFile.WriteLine "終了時刻               : " & oStpWtch.StopPoint
	objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime & " [s]"
	objLogFile.Close
	Set oStpWtch = Nothing
	Set oPrgBar = Nothing
End Function
