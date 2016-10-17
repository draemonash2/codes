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

Dim objItunes
Dim objTracks
Dim oPrgBar
Dim oStpWtch

Set objItunes = WScript.CreateObject("iTunes.Application")
Set objTracks = objItunes.LibraryPlaylist.Tracks
Set oPrgBar = New ProgressBar
Set oStpWtch = New StopWatch

Dim objLogFile
Dim sLogFilePath
Dim sDbFilePath
Dim objFso
Set objFso = CreateObject("Scripting.FileSystemObject")
sLogFilePath = sCurDir & "\" & RemoveTailWord( WScript.ScriptName, "." ) & ".log"
sDbFilePath  = sCurDir & "\" & RemoveTailWord( WScript.ScriptName, "." ) & ".db"

Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )

Call oStpWtch.StartT

' ******************************************
' * 日付入力                               *
' ******************************************
objLogFile.WriteLine "*** 日付入力処理 *** "
oPrgBar.SetMsg( _
	"⇒・日付入力処理" & vbNewLine & _
	"　・更新対象ファイル特定処理" & vbNewLine & _
	"　・Persistent ID リスト（連想配列）取得" & vbNewLine & _
	"　・タグ更新処理" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '進捗更新

On Error Resume Next
Dim sTimeValue
Dim sDateValue
Dim sCmpBaseTime

oPrgBar.SetProg( 10 ) '進捗更新

sCmpBaseTime = InputBox( _
					"更新対象とするファイルを特定します。" & vbNewLine & _
					"更新対象とする時刻を入力してください。" & vbNewLine & _
					"" & vbNewLine & _
					"  [入力規則] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
					"" & vbNewLine & _
					"※ 日付のみを指定したい場合、「YYYY/MM/DD 0:0:0」としてください。" _
					, "入力" _
					, Now() _
				)

objLogFile.WriteLine "[入力値]   " & sCmpBaseTime

sTimeValue = TimeValue(sCmpBaseTime)
sDateValue = DateValue(sCmpBaseTime)

oPrgBar.SetProg( 50 ) '進捗更新

If Err.Number <> 0 Then
	Err.Clear 'エラー情報をクリアする
	MsgBox "日付の形式が不正です！" & vbNewLine & _
	       "  [入力規則] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
	       "  [入力値] " & sCmpBaseTime
	MsgBox "プログラムを中断します！"
	Call ExecuteFinish
	WScript.Quit
End If
On Error Goto 0 '「On Error Resume Next」を解除

oPrgBar.SetProg( 100 ) '進捗更新

objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime

' ******************************************
' * 更新対象ファイル特定                   *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** 更新対象ファイル特定処理 *** "
oPrgBar.SetMsg( _
	"　・日付入力処理" & vbNewLine & _
	"⇒・更新対象ファイル特定処理" & vbNewLine & _
	"　・Persistent ID リスト（連想配列）取得" & vbNewLine & _
	"　・タグ更新処理" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '進捗更新

Dim asAllFileList()
Dim asTrgtFileList()
ReDim asAllFileList(-1)
ReDim asTrgtFileList(-1)

oPrgBar.SetProg( 20 ) '進捗更新
Call GetFileList(TRGT_DIR, asAllFileList, 1)
oPrgBar.SetProg( 20 ) '進捗更新
Call GetTrgtFileList(asAllFileList, asTrgtFileList)
'Call OutputAllElement(asTrgtFileList) '★DebugDel★

oPrgBar.SetProg( 100 ) '進捗更新

objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime

'If 0 Then ' ★DebugDel★
' ******************************************
' * Persistent ID リスト（連想配列）取得   *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** Persistent ID リスト（連想配列）取得処理 *** "
oPrgBar.SetMsg( _
	"　・日付入力処理" & vbNewLine & _
	"　・更新対象ファイル特定処理" & vbNewLine & _
	"⇒・Persistent ID リスト（連想配列）取得" & vbNewLine & _ 
	"　・タグ更新処理" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '進捗更新

Dim htPath2PerID
Set htPath2PerID = CreateObject("Scripting.Dictionary")

Dim lTrackNum
lTrackNum = objTracks.Count
'lTrackNum = 1000

Dim iTrackIdx
Dim sPersistentID
Dim sLocation

On Error Resume Next
For iTrackIdx = 1 To lTrackNum
	oPrgBar.SetProg( ( iTrackIdx / lTrackNum ) * 100 ) '進捗更新
	IF objTracks.Item( iTrackIdx ).KindAsString = "MPEG オーディオファイル" Then
		sPersistentID = GetPerIDFromObj( objTracks.Item( iTrackIdx ) )
		sLocation = objTracks.Item( iTrackIdx ).Location
		If htPath2PerID.Exists( sLocation ) = True Then
			'Do Nothing
		Else
			htPath2PerID.Add sLocation, sPersistentID
			objLogFile.WriteLine sLocation & Chr(9) & sPersistentID
		End If
		
		If Err.Number <> 0 Then
			MsgBox "Err.Number    :" & Err.Number & vbNewLine & _
				   "iTrackIdx     :" & iTrackIdx & vbNewLine & _
				   "Location      :" & Location & vbNewLine & _
				   "sPersistentID :" & sPersistentID
			Call ExecuteFinish
		Else
			'Do Nothing
		End If
	Else
		'Do Nothing
	End If
Next
objLogFile.WriteLine "ファイル数：" & htPath2PerID.Count
Err.Clear
On Error Goto 0

objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime

' ******************************************
' * タグ更新                               *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** タグ更新処理 *** "
oPrgBar.SetMsg( _
	"　・日付入力処理" & vbNewLine & _
	"　・更新対象ファイル特定処理" & vbNewLine & _
	"　・Persistent ID リスト（連想配列）取得" & vbNewLine & _ 
	"⇒・タグ更新処理" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '進捗更新

Dim lTrgtFileListIdx
Dim lTrgtFileListNum
Dim sFilePath
Dim objTrgtTrack
lTrgtFileListNum = UBound( asTrgtFileList )
For lTrgtFileListIdx = 0 To lTrgtFileListNum
	oPrgBar.SetProg( ( ( lTrgtFileListIdx + 1 ) / ( lTrgtFileListNum + 1 ) ) * 100 ) '進捗更新
	sFilePath = asTrgtFileList( lTrgtFileListIdx )
	sPersistentID = htPath2PerID.Item( sFilePath )
	
	Set objTrgtTrack = GetObjFromPerID( sPersistentID )
	objTrgtTrack.Composer = "1"
	objTrgtTrack.Composer = ""
	Set objTrgtTrack = Nothing
Next

objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime

'End If ' ★DebugDel★

' ******************************************
' * 終了処理                               *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** 終了処理 *** "

Call ExecuteFinish()

MsgBox "プログラムが正常に終了しました。"

'==========================================================
'= 関数定義
'==========================================================
Function ExecuteFinish()
	
	Call oStpWtch.StopT
	
	objLogFile.WriteLine "StartPoint      : " & oStpWtch.StartPoint
	objLogFile.WriteLine "StopPoint       : " & oStpWtch.StopPoint
	objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
	objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime
	
	objLogFile.Close
	Call oPrgBar.Quit
	
	Set oStpWtch = Nothing
	Set oPrgBar = Nothing
End Function

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

Function GetTrgtFileList( _
	ByRef asAllFileList, _
	ByRef asTrgtFileList _
)
	Dim lAllFileListIdx
	Dim lTrgtFileListIdx
	Dim sExtName
	Dim sLastModDate
	Dim oFileSys
	Dim sFilePath
	
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	
	For lAllFileListIdx = 0 to UBound(asAllFileList)
		oPrgBar.SetProg( ( lAllFileListIdx / UBound(asAllFileList) ) * 100 ) '進捗更新
		sExtName = ExtractTailWord( asAllFileList(lAllFileListIdx), "." )
		sFilePath = asAllFileList(lAllFileListIdx)
		If LCase(sExtName) = "mp3" Then
			sLastModDate = oFileSys.GetFile(sFilePath).DateLastModified
			If DateDiff("s", sCmpBaseTime, sLastModDate ) >= 0  Then
				ReDim Preserve asTrgtFileList( UBound(asTrgtFileList) + 1 )
				asTrgtFileList( UBound(asTrgtFileList) ) = sFilePath
				objLogFile.WriteLine sFilePath
			Else
				'Do Nothing
			End If
		Else
			'Do Nothing
		End If
	Next
	objLogFile.WriteLine "ファイル数：" & UBound(asAllFileList) + 1
	
	Set oFileSys = Nothing
End Function
