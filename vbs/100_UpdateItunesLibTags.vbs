Option Explicit

Const TRGT_DIR = "Z:\300_Musics"
 
Dim objItunes
Dim objTracks
Dim oPrgBar
Dim oStpWtch

Set objItunes = WScript.CreateObject("iTunes.Application")
Set objTracks = objItunes.LibraryPlaylist.Tracks
Set oPrgBar = New ProgressBar
Set oStpWtch = New StopWatch

Dim objWshShell
Dim objLogFile
Dim sLogFilePath
Dim objFso

Set objFso = CreateObject("Scripting.FileSystemObject")
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
sLogFilePath = objWshShell.CurrentDirectory & "\" & _
               RemoveTailWord( WScript.ScriptName, "." ) & "_log.txt"

Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )

Call oStpWtch.StartT

objLogFile.WriteLine "[現在時刻] " & Now()

' ******************************************
' * 日付入力                               *
' ******************************************
objLogFile.WriteLine ""
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

objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime

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

objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime

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
		sPersistentID = PersistentID( objTracks.Item( iTrackIdx ) )
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
objLogFile.WriteLine "取得完了ファイル数：" & htPath2PerID.Count
Err.Clear
On Error Goto 0

objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime

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
lTrgtFileListNum = UBound( asTrgtFileList )
For lTrgtFileListIdx = 0 To lTrgtFileListNum
	oPrgBar.SetProg( ( ( lTrgtFileListIdx + 1 ) / ( lTrgtFileListNum + 1 ) ) * 100 ) '進捗更新
	sFilePath = asTrgtFileList( lTrgtFileListIdx )
	sPersistentID = htPath2PerID.Item( sFilePath )
	
	ObjectFromID( sPersistentID ).Composer = "1"
	ObjectFromID( sPersistentID ).Composer = ""
Next

objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime

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

	objLogFile.WriteLine "StartTime    : " & oStpWtch.StartTime
	objLogFile.WriteLine "StopTime     : " & oStpWtch.StopTime
	objLogFile.WriteLine "LapTime      : " & oStpWtch.LapTime
	objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime
	
	objLogFile.Close
	Call oPrgBar.Quit

	Set oStpWtch = Nothing
	Set oPrgBar = Nothing

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
	
	Set oFileSys = Nothing
End Function

Function OutputAllElement( _
	ByRef asOutTrgtArray _
)
	Dim lIdx
	Dim sOutStr
	sOutStr = "EleNum :" & Ubound( asOutTrgtArray ) + 1
	For lIdx = 0 to UBound( asOutTrgtArray )
		sOutStr = sOutStr & vbNewLine & asOutTrgtArray(lIdx)
	Next
	WScript.Echo sOutStr
End Function

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を返却する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        抽出文字列
' = 覚書    なし
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
 
' ==================================================================
' = 概要    末尾区切り文字以降の文字列を除去する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        除去文字列
' = 覚書    なし
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

'lFileListType）0：両方、1:ファイル、2:フォルダ、それ以外：格納しない
Function GetFileList( _
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

Class ProgressBar
	Dim objExplorer
	Dim sProgressMsg
	
	Private Sub Class_Initialize()
		Dim objWMIService
		Dim colItems
		Dim strComputer
		Dim objItem
		Dim intHorizontal
		Dim intVertical
		strComputer = "."
		Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
		For Each objItem in colItems
			intHorizontal = objItem.ScreenWidth
			intVertical = objItem.ScreenHeight
		Next
		Set objWMIService = Nothing
		Set colItems = Nothing
		
		Set objExplorer = CreateObject("InternetExplorer.Application")
		objExplorer.Navigate "about:blank"
		objExplorer.ToolBar = 0
		objExplorer.StatusBar = 0
		objExplorer.Left = (intHorizontal - 400) / 2
		objExplorer.Top = (intVertical - 200) / 2
		objExplorer.Width = 400
		objExplorer.Height = 500
		objExplorer.Visible = 1
		
		Call ActiveIE
		objExplorer.Document.Body.Style.Cursor = "wait"
		objExplorer.Document.Title = "進捗状況"
		SetProg(0)
	End Sub
	
	Private Sub Class_Terminate()
		'Do Nothing
	End Sub
	
	Public Function SetMsg( _
		ByVal sMessage _
	)
		'改行文字を<br>に置換
		sProgressMsg = Replace( sMessage, vbNewLine, "<br>" )
	End Function
	
	Public Function SetProg( _
		ByVal lProgress _
	)
		Dim lProgress100
		Dim lProgress10
	
		If lProgress > 100 Or lProgress < 0 Then
			MsgBox "プログレスバーの進捗に規定値[0-100]外の値が指定されています！" & vbNewLine & _
				   "値：" & lProgress
			MsgBox "プログラムを中止します！"
			Call ProgressBarQuit
			WScript.Quit
		End If
		
		lProgress100 = Fix(lProgress)
		lProgress10 = Fix(lProgress / 10)
		
		objExplorer.Document.Body.InnerHTML = sProgressMsg & "<br>" & _
											  "<br>" & _
											  "処理中..." & "<br>" & _
											  String( lProgress10, "■") & String( 10 - lProgress10, "□") & _
											  "  " & lProgress100 & "% 完了"
	End Function
	
	Public Function Quit()
		objExplorer.Document.Body.Style.Cursor = "default"
		objExplorer.Quit
	End Function
	
	Private Function ActiveIE()
		Dim objWshShell
		Dim intProcID
	
		Const strIEexe = "iexplore.exe" 'IEのプロセス名
		intProcID = GetProcID(strIEexe)
		Set objWshShell = CreateObject("Wscript.Shell")
		objWshShell.AppActivate intProcID
		Set objWshShell = Nothing
	End Function
	
	Private Function GetProcID(ProcessName)
		Dim Service
		Dim QfeSet
		Dim Qfe
		Dim intProcID
		
		Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
		Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='"& ProcessName &"'")
		
		intProcID = 0
		
		For Each Qfe in QfeSet
			intProcID = Qfe.ProcessId
			GetProcID = intProcID
			Exit For
		Next
	End Function
End Class

Function ObjectFromID( sID )
	Set ObjectFromID = objItunes.LibraryPlaylist.Tracks.ItemByPersistentID(Eval("&H" & Left( sID, 8)),Eval("&H" & Right( sID, 8)))
End Function

Function PersistentID( objT )
	PersistentID = Right("0000000" & Hex(objItunes.ITObjectPersistentIDHigh(objT)),8) & _
	               Right("0000000" & Hex(objItunes.ITObjectPersistentIDLow(objT)),8)
End Function

Class StopWatch
	Dim glStartTime
	Dim glStopTime
	Dim glIntervalTime
	
	Private Sub Class_Initialize()
		Call StopWatchInit
	End Sub
	
	'*** 初期化 ***
	Private Function StopWatchInit()
		glStartTime = 0
		glStopTime = 0
		glIntervalTime = 0
	End Function
	
	'*** 測定開始 ***
	Public Sub StartT()
		glStartTime = Now()
		glIntervalTime = glStartTime
	End Sub
	
	'*** 測定停止 ***
	Public Function StopT()
		glStopTime = Now()
		StopT = glStopTime - glStartTime
	End Function
	
	'*** 開始から現在までの総経過時間[s] ***
	Public Property Get LapTime()
'		LapTime = Now() - glStartTime
		LapTime = DateDiff( "s", glStartTime, Now() )
	End Property
	
	'*** 前回 IntervalTime() 呼び出し時からの時間間隔[s] ***
	Public Property Get IntervalTime()
		Dim lCurTime
		lCurTime = Now()
'		IntervalTime = lCurTime - glIntervalTime
		IntervalTime = DateDiff( "s", glIntervalTime, lCurTime )
		glIntervalTime = lCurTime
	End Property
	
	'*** 開始時刻 ***
	Public Property Get StartTime()
		StartTime = glStartTime
	End Property
	
	'*** 終了時刻 ***
	Public Property Get StopTime()
		StopTime = glStopTime
	End Property
	
	Private Sub Class_Terminate()
		Call StopWatchInit
	End Sub
End Class
