Option Explicit

'□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□
'□
'□ iTunes タグ更新ツール
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
'□     ・指定したフォルダ配下に実行結果ログを格納する。
'□     ・「作曲者」「グループ」「コメント」（他にもあるかも）においては iTunes 以外の
'□       ソフトで更新しても、iTunes 上の表示は更新されない。（iTunes の仕様） 
'□       表示を更新するためには、iTunes から直接上記タグを更新する必要がある。
'□       本ツールも同様の理由で、上記タグの表示更新はできない。
'□     
'□  【使用方法】
'□     (1) 音楽を格納するフォルダのルートフォルダパスを「TRGT_DIR」に記載する。
'□     (2) 「TRGT_DIR」に記載する。
'□     (3) 本スクリプトを実行。
'□     
'□  【更新履歴】
'□     v2.8 (2017/03/21)
'□       ・書き換え対象タグを変更
'□           Composer ⇒ VolumeAdjustment
'□       ・更新日書き戻し処理変更
'□     
'□     v2.7 (2017/03/15)
'□       ・古い iTunes Library Backup フォルダの削除
'□     
'□     v2.6 (2017/03/05)
'□       ・StopWatch.vbs 修正に対する対応
'□       ・軽微なバグFix
'□     
'□     v2.5 (2016/10/28)
'□       ・ログファイル出力先変更
'□       ・処理終了時にログファイルを開く処理を追加
'□     
'□     v2.4 (2016/10/27)
'□       ・処理選択追加
'□     
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
Const UPDATE_MODIFIED_DATE = False
Const ITUNES_BACKUP_FOLDER_MAX = 20

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
Call Include( "C:\codes\vbs\_lib\String.vbs" )          'GetDirPath()
                                                        'RemoveTailWord()
                                                        'ExtractTailWord()
Call Include( "C:\codes\vbs\_lib\StopWatch.vbs" )       'class StopWatch
Call Include( "C:\codes\vbs\_lib\ProgressBarIE.vbs" )   'class ProgressBar
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )      'GetFileList2()
Call Include( "C:\codes\vbs\_lib\iTunes.vbs" )          '★←インクルードいる？
Call Include( "C:\codes\vbs\_lib\Array.vbs" )           '★←インクルードいる？

' ******************************************
' * 処理選択                               *
' ******************************************
Dim bIsExecLibAdd
Dim bIsExecLibMod
Dim sAnswer
sAnswer = MsgBox( "iTunes へファイルを追加しますか？" & vbNewLine & _
                  "  [追加対象フォルダ] " & TRGT_DIR _
                  , vbYesNoCancel _
                )
Select Case sAnswer
    Case vbYes: bIsExecLibAdd = True
    Case vbNo:  bIsExecLibAdd = False
    Case Else:
        MsgBox "処理をキャンセルしました。" & vbNewLine & _
               "プログラムを終了します。"
        WScript.Quit
End Select
sAnswer = MsgBox( "登録済みの曲について iTunes ライブラリのタグを更新しますか？" & vbNewLine & _
                  "  [更新対象フォルダ] " & TRGT_DIR _
                  , vbYesNoCancel _
                )
Select Case sAnswer
    Case vbYes: bIsExecLibMod = True
    Case vbNo:  bIsExecLibMod = False
    Case Else:
        MsgBox "処理をキャンセルしました。" & vbNewLine & _
               "プログラムを終了します。"
        WScript.Quit
End Select
sAnswer = MsgBox( "実行する処理は以下で問題ありませんか？" & vbNewLine & _
                  "  ・iTunes ライブラリ追加         ： " & bIsExecLibAdd & vbNewLine & _
                  "  ・iTunes ライブラリ更新         ： " & bIsExecLibMod _
                  , vbYesNo _
                )
Select Case sAnswer
    Case vbYes:
        'Do Nothing
    Case vbNo:
        MsgBox "処理をキャンセルしました。" & vbNewLine & _
               "プログラムを終了します。"
        WScript.Quit
End Select

' ******************************************
' * 事前処理                               *
' ******************************************
If 1 Then '処理ブロック化のための分岐処理
    '*** ストップウォッチ起動 ***
    Dim oStpWtch
    Set oStpWtch = New StopWatch
    Call oStpWtch.StartT
    
    '*** プログレスバー起動 ***
    Dim oPrgBar
    Set oPrgBar = New ProgressBar
End If

' ******************************************
' * iTunes ライブラリバックアップ          *
' ******************************************
If DEBUG_FUNCVALID_BACKUPITUNELIBRARYS = True Then ' ★Debug★
    oPrgBar.Message = _
        "⇒・iTunes ライブラリ バックアップ" & vbNewLine & _
        "　・iTunes ライブラリ 追加処理" & vbNewLine & _
        "　・iTunes ライブラリ 更新処理" & vbNewLine & _
        "　　- 日付入力処理" & vbNewLine & _
        "　　- 更新対象ファイル特定処理" & vbNewLine & _
        "　　- タグ更新処理" & _
        ""
    oPrgBar.Update( 0.2 ) '進捗更新
    
    '*** バックアップフォルダ作成 ***
    Dim sCurDateTime
    sCurDateTime = Now()
    
    Dim objItunes
    Set objItunes = WScript.CreateObject("iTunes.Application")
    
    Dim sBackUpDirName
    Dim sBackUpDirPath
    Dim sItuneDirPath
    Dim sItuneBackUpDirPath
    sItuneDirPath = GetDirPath( objItunes.LibraryXMLPath )
    sItuneBackUpDirPath = sItuneDirPath & "\iTunes Library Backup"
    sBackUpDirName = Year( sCurDateTime ) & _
                     String( 2 - Len( Month( sCurDateTime ) ), "0" ) & Month( sCurDateTime ) & _
                     String( 2 - Len( Day( sCurDateTime ) ), "0" ) & Day( sCurDateTime ) & _
                     "_" & _
                     String( 2 - Len( Hour( sCurDateTime ) ), "0" ) & Hour( sCurDateTime ) & _
                     String( 2 - Len( Minute( sCurDateTime ) ), "0" ) & Minute( sCurDateTime ) & _
                     String( 2 - Len( Second( sCurDateTime ) ), "0" ) & Second( sCurDateTime )
    sBackUpDirPath = sItuneBackUpDirPath & "\" & sBackUpDirName
    
    Set objItunes = Nothing
    
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder( sBackUpDirPath )
    objFSO.CopyFile sItuneDirPath & "\iTunes Library Extras.itdb", sBackUpDirPath & "\"
    objFSO.CopyFile sItuneDirPath & "\iTunes Library Genius.itdb", sBackUpDirPath & "\"
    objFSO.CopyFile sItuneDirPath & "\iTunes Library.itl        ", sBackUpDirPath & "\"
    objFSO.CopyFile sItuneDirPath & "\iTunes Music Library.xml  ", sBackUpDirPath & "\"
    
    '*** 古いバックアップフォルダを削除 ***
    Dim asDirList
    Call GetFileList2(sItuneBackUpDirPath, asDirList, 2) 
    
    'フォルダ削除
    If UBound( asDirList ) >= ITUNES_BACKUP_FOLDER_MAX then
        Dim lDelFolderMax
        lDelFolderMax = UBound(asDirList) - ITUNES_BACKUP_FOLDER_MAX
        Dim lDelDirIdx
        For lDelDirIdx = LBound(asDirList) to lDelFolderMax
            'バックアップフォルダ名は「YYYYMMDD_HHMMSS」で統一されているため、
            'asDirList() は自然と日時順に並ぶ。（要素番号が大きくなるほど新しい）
            'そのため、要素番号の小さい順からフォルダを削除する。
            objFSO.DeleteFolder asDirList(lDelDirIdx), True
        Next
    Else
        'Do Nothing
    End If
    
    '*** ログファイル作成 ***
    Dim sLogFilePath
    sLogFilePath = sBackUpDirPath & "\" & Replace( WScript.ScriptName, ".vbs", ".log" )
    
    Dim objLogFile
    Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )
    
    objLogFile.WriteLine "script started."
    objLogFile.WriteLine ""
    objLogFile.WriteLine "[更新対象フォルダ] " & TRGT_DIR
    objLogFile.WriteLine ""
    objLogFile.WriteLine "*** iTunes ライブラリバックアップ *** "
    objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime
    objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime
Else
    sLogFilePath = TRGT_DIR & "\" & Replace( WScript.ScriptName, ".vbs", ".log" )
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )
    
    objLogFile.WriteLine "script started."
    objLogFile.WriteLine ""
    objLogFile.WriteLine "[更新対象フォルダ] " & TRGT_DIR
End If ' ★Debug★

' ******************************************
' * ファイル追加                           *
' ******************************************
if bIsExecLibAdd = True Then
    
    If DEBUG_FUNCVALID_ADDFILES = True Then ' ★Debug★
        
        objLogFile.WriteLine ""
        objLogFile.WriteLine "*** ファイル追加 *** "
        oPrgBar.Message = _
            "　・iTunes ライブラリ バックアップ" & vbNewLine & _
            "⇒・iTunes ライブラリ 追加処理" & vbNewLine & _
            "　・iTunes ライブラリ 更新処理" & vbNewLine & _
            "　　- 日付入力処理" & vbNewLine & _
            "　　- 更新対象ファイル特定処理" & vbNewLine & _
            "　　- タグ更新処理" & _
            ""
        oPrgBar.Update( 0.5 ) '進捗更新
        
        WScript.CreateObject("iTunes.Application").LibraryPlaylist.AddFile( TRGT_DIR )
        
        oPrgBar.Update( 1 ) '進捗更新
        
        objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime
        objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime
    Else
        'Do Nothing
    End If ' ★Debug★
Else
    'Do Nothing
End If

' ******************************************
' * ライブラリ更新                         *
' ******************************************
if bIsExecLibMod = True Then
    ' ==============================
    ' = 日付入力                   =
    ' ==============================
    If DEBUG_FUNCVALID_DATEINPUT = True Then ' ★Debug★
        
        objLogFile.WriteLine ""
        objLogFile.WriteLine "*** 日付入力処理 *** "
        oPrgBar.Message = _
            "　・iTunes ライブラリ バックアップ" & vbNewLine & _
            "　・iTunes ライブラリ 追加処理" & vbNewLine & _
            "　・iTunes ライブラリ 更新処理" & vbNewLine & _
            "⇒　- 日付入力処理" & vbNewLine & _
            "　　- 更新対象ファイル特定処理" & vbNewLine & _
            "　　- タグ更新処理" & _
            ""
        oPrgBar.Update( 0 ) '進捗更新
        
        On Error Resume Next
        
        oPrgBar.Update( 0.1 ) '進捗更新
        
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
        
        oPrgBar.Update( 0.5 ) '進捗更新
        
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
        
        oPrgBar.Update( 1 ) '進捗更新
        
        oStpWtch.IntervalTime ' IntervalTime 更新
        
    Else ' ★Debug★
        sCmpBaseTime = "2016/10/27 10:00:00"
    End If ' ★Debug★
    
    ' ==============================
    ' = 更新対象ファイルリスト取得 =
    ' ==============================
    If DEBUG_FUNCVALID_TRGTLISTUP = True Then ' ★Debug★
        
        objLogFile.WriteLine ""
        objLogFile.WriteLine "*** 更新対象ファイル特定 *** "
        oPrgBar.Message = _
            "　・iTunes ライブラリ バックアップ" & vbNewLine & _
            "　・iTunes ライブラリ 追加処理" & vbNewLine & _
            "　・iTunes ライブラリ 更新処理" & vbNewLine & _
            "　　- 日付入力処理" & vbNewLine & _
            "⇒　- 更新対象ファイル特定処理" & vbNewLine & _
            "　　- タグ更新処理" & _
            ""
        oPrgBar.Update( 0 ) '進捗更新
        
        On Error Resume Next
        
        '*** Dir コマンド実行 ***
        Dim sTmpFilePath
        Dim sExecCmd
        sTmpFilePath = objWshShell.CurrentDirectory & "\" & replace( WScript.ScriptName, ".vbs", "_TrgtFileList.tmp" )
        If DEBUG_FUNCVALID_DIRCMDEXEC = True Then ' ★Debug★
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
        
        oPrgBar.Update( 0.2 ) '進捗更新
        
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
                oPrgBar.Update( _
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
        
        oPrgBar.Update( 1 ) '進捗更新
        
        If DEBUG_FUNCVALID_DIRRESULTDELETE = True Then ' ★Debug★
            objFSO.DeleteFile sTmpFilePath, True
        End If ' ★Debug★
        
        Set objFSO = Nothing    'オブジェクトの破棄
        
        objLogFile.WriteLine "ファイル数：" & UBound(asTrgtFileList) + 1
        objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime
        objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime
        
    Else ' ★Debug★
        ReDim asTrgtFileList(0)
        asTrgtFileList(0) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Bow Down.mp3"
    '   asTrgtFileList(1) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concentrate.mp3"
    '   asTrgtFileList(2) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concrete Schoolyard.mp3"
    '   asTrgtFileList(3) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Control Myself.mp3"
    End If ' ★Debug★
    
    ' ==============================
    ' = タグ更新                   =
    ' ==============================
    If DEBUG_FUNCVALID_TAGUPDATE = True Then ' ★Debug★
        
        objLogFile.WriteLine ""
        objLogFile.WriteLine "*** タグ更新処理 *** "
        oPrgBar.Message = _
            "　・iTunes ライブラリ バックアップ" & vbNewLine & _
            "　・iTunes ライブラリ 追加処理" & vbNewLine & _
            "　・iTunes ライブラリ 更新処理" & vbNewLine & _
            "　　- 日付入力処理" & vbNewLine & _
            "　　- 更新対象ファイル特定処理" & vbNewLine & _
            "⇒　- タグ更新処理" & _
            ""
        oPrgBar.Update( 0 ) '進捗更新
        
        objLogFile.WriteLine "[TrgtFileIdx}" & Chr(9) & "[HitIdx] / [HitNum}" & Chr(9) & "[FilePath}" & Chr(9) & "[TrackName]" & Chr(9) & "[Kind]" & Chr(9) & "[Location]" & Chr(9) & "[LocMatch]"
        
        Dim lTrgtFileListIdx
        Dim lTrgtFileListNum
        lTrgtFileListNum = UBound( asTrgtFileList )
        For lTrgtFileListIdx = 0 To lTrgtFileListNum
            '進捗更新
            oPrgBar.Update( _
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
            
            Dim sOutLine
            Dim bIsFilePathMatched
            Dim lHitIdx
            bIsFilePathMatched = False
            For lHitIdx = 1 to objSearchResult.Count
                With objSearchResult.Item(lHitIdx)
                    sOutLine = ( lTrgtFileListIdx + 1 ) & Chr(9) & lHitIdx & " / " & objSearchResult.Count & Chr(9) & sTrgtFilePath & Chr(9) & sTrgtTrackName & Chr(9) & .Kind
                    If .Kind = 1 Then
                        sOutLine = sOutLine & Chr(9) & .Location & Chr(9) & ( .Location = sTrgtFilePath )
                        If .Location = sTrgtFilePath Then
                            .VolumeAdjustment = 0
                            bIsFilePathMatched = True
                        Else
                            'Do Nothing
                        End If
                    Else
                        'Do Nothing
                    End If
                End With
                objLogFile.WriteLine sOutLine
                If bIsFilePathMatched = True Then
                    Exit For
                Else
                    'Do Nothing
                End If
            Next
            Set objSearchResult = Nothing
            Set objPlayList = Nothing
            
            If UPDATE_MODIFIED_DATE = True Then
                '更新日は更新されたまま
            Else
                oTrgtFile.ModifyDate = CDate( sTrgtModDate ) '更新日を書き戻し
            End If
            
            Set oTrgtFile = Nothing
            Set oTrgtDirFiles = Nothing
        Next
        
        objLogFile.WriteLine "ファイル数：" & UBound(asTrgtFileList) + 1
        objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime
        objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime
        
    End If ' ★Debug★
Else
    'Do Nothing
End If

' ******************************************
' * 終了処理                               *
' ******************************************
Call Finish
WScript.CreateObject("WScript.Shell").Run "%comspec% /c """ & sLogFilePath & """"
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
    objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime
    objLogFile.WriteLine ""
    objLogFile.WriteLine "script finished."
    objLogFile.Close
    Set oStpWtch = Nothing
    Set oPrgBar = Nothing
End Function
