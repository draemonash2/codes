Option Explicit

'□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□
'□
'□ iTunes タグ更新ツール v.2.3
'□
'□  【概要】
'□     mp3 ファイルの更新日時を元にタグ更新済みファイルを自動判別し、iTunes の API を
'□     コマンドラインから実行して、iTunes ライブラリの mp3 タグ情報を更新する。
'□     
'□     iTunes 以外のソフトで mp3 タグを更新した場合、iTunes を起動しても古いタグの
'□     まま表示され、ライブラリは更新されない。この時、該当のファイルを再生するか、
'□     iTunes 上で何らかのタグを更新すればライブラリ更新されるが、大量のファイルを
'□     更新した場合、非常に手間がかかる。本ツールは、その手間を省くことを目的として
'□     作成した。
'□     
'□     また、指定フォルダ配下のライブラリ未登録のファイルも自動的に登録できる。
'□     （登録実行可否は選択可）
'□     
'□  【注意事項、特記事項】
'□     ・本ツールは、mp3 タグ内の未使用のエントリを更新することで、ライブラリを
'□     更新する。この時、エントリ「作曲者(Composer)」を書き換えるため、
'□     ここに何らかの情報を埋め込んでいる場合、空白になってしまう。
'□     ・指定したフォルダ配下に実行結果ログを格納する。
'□     
'□  【使用方法】
'□     (1) 音楽を格納するフォルダのルートフォルダパスを「TRGT_DIR」に記載する。
'□     (2) 「TRGT_DIR」に記載する。
'□     (3) 本スクリプトを実行。
'□     
'□  【更新履歴】
'□     v2.3 (2016/10/18)
'□       ・iTunes ライブラリバックアップ機能追加
'□     
'□     v2.2 (2016/10/17)
'□       ・初版
'□     
'□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□

'==========================================================
'= 設定値
'==========================================================
Const TRGT_DIR = "Z:\300_Musics"

Const DEBUG_FUNCVALID_DISABLEUPDATEMODDATE  = True
Const DEBUG_FUNCVALID_BACKUPITUNELIBRARYS   = True
Const DEBUG_FUNCVALID_ADDFILES              = True
Const DEBUG_FUNCVALID_DATEINPUT             = True
Const DEBUG_FUNCVALID_TRGTLISTUP            = True
Const DEBUG_FUNCVALID_DIRCMDEXEC            = True
Const DEBUG_FUNCVALID_TAGUPDATE             = True
Const DEBUG_FUNCVALID_DIRRESULTDELETE       = True

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
sLogFilePath = TRGT_DIR & "\" & RemoveTailWord( WScript.ScriptName, "." ) & ".log"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objLogFile
Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )

objLogFile.WriteLine "script started."
objLogFile.WriteLine ""
objLogFile.WriteLine "[更新対象フォルダ] " & TRGT_DIR

Dim oStpWtch
Set oStpWtch = New StopWatch
Call oStpWtch.StartT

Dim oPrgBar
Set oPrgBar = New ProgressBar

' ******************************************
' * iTunes ライブラリバックアップ          *
' ******************************************
If DEBUG_FUNCVALID_BACKUPITUNELIBRARYS Then ' ★Debug★

objLogFile.WriteLine ""
objLogFile.WriteLine "*** iTunes ライブラリバックアップ *** "
oPrgBar.SetMsg( _
    "⇒・iTunes ライブラリバックアップ" & vbNewLine & _
    "　・ファイル追加処理" & vbNewLine & _
    "　・日付入力処理" & vbNewLine & _
    "　・更新対象ファイル特定処理" & vbNewLine & _
    "　・タグ更新処理" & _
    "" _
)
oPrgBar.SetProg( 20 ) '進捗更新

Dim sCurDateTime
sCurDateTime = Now()

Dim objItunes
Set objItunes = WScript.CreateObject("iTunes.Application")

Dim sBackUpDirName
Dim sBackUpDirPath
Dim sItuneDirPath
sItuneDirPath = GetDirPath( objItunes.LibraryXMLPath )
sBackUpDirName = Year( sCurDateTime ) & _
                 String( 2 - Len( Month( sCurDateTime ) ), "0" ) & Month( sCurDateTime ) & _
                 String( 2 - Len( Day( sCurDateTime ) ), "0" ) & Day( sCurDateTime ) & _
                 "_" & _
                 String( 2 - Len( Hour( sCurDateTime ) ), "0" ) & Hour( sCurDateTime ) & _
                 String( 2 - Len( Minute( sCurDateTime ) ), "0" ) & Minute( sCurDateTime ) & _
                 String( 2 - Len( Second( sCurDateTime ) ), "0" ) & Second( sCurDateTime )
sBackUpDirPath = sItuneDirPath & "\iTunes Library Backup\" & sBackUpDirName

Set objItunes = Nothing

objFSO.CreateFolder( sBackUpDirPath )
objFSO.CopyFile sItuneDirPath & "\iTunes Library Extras.itdb", sBackUpDirPath & "\"
objFSO.CopyFile sItuneDirPath & "\iTunes Library Genius.itdb", sBackUpDirPath & "\"
objFSO.CopyFile sItuneDirPath & "\iTunes Library.itl        ", sBackUpDirPath & "\"
objFSO.CopyFile sItuneDirPath & "\iTunes Music Library.xml  ", sBackUpDirPath & "\"

oPrgBar.SetProg( 100 ) '進捗更新

objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime & " [s]"

End If ' ★Debug★

' ******************************************
' * ファイル追加                           *
' ******************************************
If DEBUG_FUNCVALID_ADDFILES Then ' ★Debug★

objLogFile.WriteLine ""
objLogFile.WriteLine "*** ファイル追加 *** "
oPrgBar.SetMsg( _
    "　・iTunes ライブラリバックアップ" & vbNewLine & _
    "⇒・ファイル追加処理" & vbNewLine & _
    "　・日付入力処理" & vbNewLine & _
    "　・更新対象ファイル特定処理" & vbNewLine & _
    "　・タグ更新処理" & _
    "" _
)
oPrgBar.SetProg( 50 ) '進捗更新

Dim sAnswer
sAnswer = MsgBox( "iTunes へファイルを追加しますか？" & vbNewLine & _
                  "  [追加対象フォルダ] " & TRGT_DIR _
                  , vbYesNoCancel _
                )
If sAnswer = vbYes Then
    MsgBox "iTunes へ " & TRGT_DIR & " を追加します。"
    WScript.CreateObject("iTunes.Application").LibraryPlaylist.AddFile( TRGT_DIR )
ElseIf sAnswer = vbNo Then
    MsgBox "iTunes への追加をスキップします。"
Else
    MsgBox "プログラムを中断します。"
    Call Finish
    WScript.Quit
End If
oPrgBar.SetProg( 100 ) '進捗更新

objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime & " [s]"
End If ' ★Debug★

sAnswer = MsgBox( "登録済みの曲について iTunes ライブラリのタグを更新しますか？" & vbNewLine & _
                  "  [更新対象フォルダ] " & TRGT_DIR _
                  , vbYesNoCancel _
                )
If sAnswer = vbYes Then
    MsgBox TRGT_DIR & " の iTunes ライブラリを更新します。"
ElseIf sAnswer = vbNo Then
    MsgBox "iTunes ライブラリを更新しません。"
    MsgBox "プログラムを終了します。"
    Call Finish
    WScript.Quit
Else
    MsgBox "プログラムを中断します。"
    Call Finish
    WScript.Quit
End If

' ******************************************
' * 日付入力                               *
' ******************************************
If DEBUG_FUNCVALID_DATEINPUT Then ' ★Debug★

objLogFile.WriteLine ""
objLogFile.WriteLine "*** 日付入力処理 *** "
oPrgBar.SetMsg( _
    "　・iTunes ライブラリバックアップ" & vbNewLine & _
    "　・ファイル追加処理" & vbNewLine & _
    "⇒・日付入力処理" & vbNewLine & _
    "　・更新対象ファイル特定処理" & vbNewLine & _
    "　・タグ更新処理" & _
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

oStpWtch.IntervalTime ' IntervalTime 更新

Else ' ★Debug★
sCmpBaseTime = "2016/10/16 22:50:00"
End If ' ★Debug★

' ******************************************
' * 更新対象ファイルリスト取得             *
' ******************************************
If DEBUG_FUNCVALID_TRGTLISTUP Then ' ★Debug★

objLogFile.WriteLine ""
objLogFile.WriteLine "*** 更新対象ファイル特定 *** "
oPrgBar.SetMsg( _
    "　・iTunes ライブラリバックアップ" & vbNewLine & _
    "　・ファイル追加処理" & vbNewLine & _
    "　・日付入力処理" & vbNewLine & _
    "⇒・更新対象ファイル特定処理" & vbNewLine & _
    "　・タグ更新処理" & _
    "" _
)
oPrgBar.SetProg( 0 ) '進捗更新

On Error Resume Next

'*** Dir コマンド実行 ***
Dim sTmpFilePath
Dim sExecCmd
sTmpFilePath = objWshShell.CurrentDirectory & "\" & replace( WScript.ScriptName, ".vbs", "_TrgtFileList.tmp" )
If DEBUG_FUNCVALID_DIRCMDEXEC Then ' ★Debug★
sExecCmd = "Dir """ & TRGT_DIR & """ /s /a:a-d > """ & sTmpFilePath & """"
With CreateObject("Wscript.Shell")  
    .Run "cmd /c" & sExecCmd, 7, True
End With
End If ' ★Debug★

'*** Dir コマンド結果取得 ***
Dim objFile
Dim sTextAll
If Err.Number = 0 Then
    Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
    If Err.Number = 0 Then
        sTextAll = objFile.ReadAll
        sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '末尾に改行が付与されてしまうため、削除
        objFile.Close
    Else
        WScript.Echo "ファイルが開けません: " & Err.Description
    End If
    Set objFile = Nothing   'オブジェクトの破棄
Else
    WScript.Echo "エラー " & Err.Description
End If
On Error Goto 0

oPrgBar.SetProg( 20 ) '進捗更新

'*** 更新日時抽出 ***
Dim oMatchResult
Dim sSearchPattern
Dim oRegExp
Dim sTargetStr
Set oRegExp = CreateObject("VBScript.RegExp")
sSearchPattern = "((\d{4}/\d{1,2}/\d{1,2})\s+(\d{1,2}:\d{1,2})\s+([0-9,]+)\s+(.+)\r)|(\s+(.*)\sのディレクトリ)"
sTargetStr = sTextAll
oRegExp.Pattern = sSearchPattern               '検索パターンを設定
oRegExp.IgnoreCase = True                      '大文字と小文字を区別しない
oRegExp.Global = True                          '文字列全体を検索
Set oMatchResult = oRegExp.Execute(sTargetStr) 'パターンマッチ実行

Dim sFileName
Dim sFilePath
Dim sFileSize
Dim sModDate
Dim sDirName
Dim iMatchIdx
Dim sExtName
Dim asTrgtFileList()
ReDim asTrgtFileList(-1)
Dim lMatchResultCount
sDirName = ""
objLogFile.WriteLine "[sFilePath]" & chr(9) & _
                     "[sDirName]"  & chr(9) & _
                     "[sFileName]" & chr(9) & _
                     "[sModDate]"  & chr(9) & _
                     "[sFileSize]" & chr(9) & _
                     "[sExtName]"
lMatchResultCount = oMatchResult.Count
For iMatchIdx = 0 To lMatchResultCount - 1
    '進捗更新
    If iMatchIdx Mod 100 = 0 Then
        oPrgBar.SetProg( _
            oPrgBar.ConvProgRange( _
                0, _
                oMatchResult.Count - 1, _
                iMatchIdx _
            ) _
        )
    End If
    If oMatchResult(iMatchIdx).SubMatches.Count = 7 Then
        'ファイル名にマッチ
        If oMatchResult(iMatchIdx).SubMatches(0) <> "" Then
            sModDate = oMatchResult(iMatchIdx).SubMatches(1) & " " & _
                       oMatchResult(iMatchIdx).SubMatches(2)
            sFileSize = oMatchResult(iMatchIdx).SubMatches(3)
            sFileName = oMatchResult(iMatchIdx).SubMatches(4)
            sFilePath = sDirName & "\" & sFileName
            sExtName = ExtractTailWord( sFileName, "." )
            
'           objLogFile.WriteLine oMatchResult(iMatchIdx).SubMatches(0) & chr(9) & _
'                                oMatchResult(iMatchIdx).SubMatches(1) & chr(9) & _
'                                oMatchResult(iMatchIdx).SubMatches(2) & chr(9) & _
'                                oMatchResult(iMatchIdx).SubMatches(3) & chr(9) & _
'                                oMatchResult(iMatchIdx).SubMatches(4) & chr(9) & _
'                                oMatchResult(iMatchIdx).SubMatches(5) & chr(9) & _
'                                oMatchResult(iMatchIdx).SubMatches(6)
            
            '更新日時比較 ＆ 更新対象選定
            If LCase(sExtName) = "mp3" Then
                If DateDiff("s", sCmpBaseTime, sModDate ) >= 0  Then
                    ReDim Preserve asTrgtFileList( UBound(asTrgtFileList) + 1 )
                    asTrgtFileList( UBound(asTrgtFileList) ) = sFilePath
                    objLogFile.WriteLine sFilePath & chr(9) & _
                                         sDirName  & chr(9) & _
                                         sFileName & chr(9) & _
                                         sModDate  & chr(9) & _
                                         sFileSize & chr(9) & _
                                         sExtName
                Else
                    'Do Nothing
                End If
            Else
                'Do Nothing
            End If
            
        'フォルダ名にマッチ
        ElseIf oMatchResult(iMatchIdx).SubMatches(5) <> "" Then
            sDirName = oMatchResult(iMatchIdx).SubMatches(6)
        Else
            MsgBox "エラー！"
        End If
    Else
        MsgBox "エラー！"
    End If
Next

oPrgBar.SetProg( 100 ) '進捗更新

If DEBUG_FUNCVALID_DIRRESULTDELETE Then ' ★Debug★
objFSO.DeleteFile sTmpFilePath, True
End If ' ★Debug★

Set objFSO = Nothing    'オブジェクトの破棄

objLogFile.WriteLine "ファイル数：" & UBound(asTrgtFileList) + 1
objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime & " [s]"

Else ' ★Debug★
    ReDim asTrgtFileList(0)
    asTrgtFileList(0) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Bow Down.mp3"
'   asTrgtFileList(1) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concentrate.mp3"
'   asTrgtFileList(2) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concrete Schoolyard.mp3"
'   asTrgtFileList(3) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Control Myself.mp3"
End If ' ★Debug★

' ******************************************
' * タグ更新                               *
' ******************************************
If DEBUG_FUNCVALID_TAGUPDATE Then ' ★Debug★
    
objLogFile.WriteLine ""
objLogFile.WriteLine "*** タグ更新処理 *** "
oPrgBar.SetMsg( _
    "　・iTunes ライブラリバックアップ" & vbNewLine & _
    "　・ファイル追加処理" & vbNewLine & _
    "　・日付入力処理" & vbNewLine & _
    "　・更新対象ファイル特定処理" & vbNewLine & _
    "⇒・タグ更新処理" & _
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
    
    If DEBUG_FUNCVALID_DISABLEUPDATEMODDATE = True Then
        oTrgtFile.ModifyDate = CDate( sTrgtModDate ) '更新日を書き戻し
    Else
        '更新日は更新されたまま
    End If
    
    Set oTrgtFile = Nothing
    Set oTrgtDirFiles = Nothing
Next

objLogFile.WriteLine "ファイル数：" & UBound(asTrgtFileList) + 1
objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime & " [s]"

End If ' ★Debug★

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
    objLogFile.WriteLine ""
    objLogFile.WriteLine "script finished."
    objLogFile.Close
    Set oStpWtch = Nothing
    Set oPrgBar = Nothing
End Function
