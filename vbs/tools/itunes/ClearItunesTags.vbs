Option Explicit

'==========================================================
'= 設定値
'==========================================================
Const TRGT_DIR = "Z:\300_Musics"

Const DEBUG_FUNCVALID_BACKUPITUNELIBRARYS   = True
Const DEBUG_FUNCVALID_TAGUPDATE             = True

Const UPDATE_MODIFIED_DATE = False
Const ITUNES_BACKUP_FOLDER_MAX = 20

'==========================================================
'= 本処理
'==========================================================
Dim objWshShell
Dim sCurDir
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
sCurDir = objWshShell.CurrentDirectory
Call Include( "C:\codes\vbs\_lib\String.vbs" )          'GetDirPath()
                                                        'GetFileName()
Call Include( "C:\codes\vbs\_lib\StopWatch.vbs" )       'class StopWatch
Call Include( "C:\codes\vbs\_lib\ProgressBarIE.vbs" )   'class ProgressBar
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )      'GetFileList2()
Call Include( "C:\codes\vbs\_lib\iTunes.vbs" )          '★←インクルードいる？
Call Include( "C:\codes\vbs\_lib\Array.vbs" )           '★←インクルードいる？

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
        "　・iTunes ライブラリ 更新処理" & vbNewLine & _
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
' * ライブラリ更新                         *
' ******************************************
If DEBUG_FUNCVALID_TAGUPDATE = True Then ' ★Debug★
    
    objLogFile.WriteLine ""
    objLogFile.WriteLine "*** タグ更新処理 *** "
    oPrgBar.Message = _
        "　・iTunes ライブラリ バックアップ" & vbNewLine & _
        "⇒・iTunes ライブラリ 更新処理" & vbNewLine & _
        ""
    oPrgBar.Update( 0 ) '進捗更新
    
    Dim oTrgtDirFiles
    Dim oTrgtFile
    Dim objPlayList
    Set objPlayList = WScript.CreateObject("iTunes.Application").Sources.Item(1).Playlists.ItemByName("ミュージック")
    
    objLogFile.WriteLine "[Location]" & chr(9) & "[Grouping]" & chr(9) & "[Composer]"
    
    Dim lTrgtFileListNum
    Dim lTrgtFileListIdx
    lTrgtFileListNum = objPlayList.Tracks.Count
    lTrgtFileListIdx = 1
    
    Dim lHitFileNum
    lHitFileNum = 1
    
    Dim objTrack
    For Each objTrack In objPlayList.Tracks
        '進捗更新
        oPrgBar.Update( _
            oPrgBar.ConvProgRange( _
                1, _
                lTrgtFileListNum, _
                lTrgtFileListIdx _
            ) _
        )
        
        If objTrack.KindAsString = "MPEG オーディオファイル" Then
            'If InStr( objTrack.Location, "Z:\300_Musics\290_Reggae@Riddim\Cardiac Keys" ) > 0 Then '★
            If objTrack.Grouping = "" And objTrack.Composer = "" Then
                'Do Nothing
            Else
                '更新日退避
                If UPDATE_MODIFIED_DATE = True Then
                    '更新日は更新されたまま
                Else
                    Dim sTrgtFilePath
                    Dim sTrgtDirPath
                    Dim sTrgtFileName
                    Dim sTrgtModDate
                    sTrgtFilePath = objTrack.Location
                    sTrgtDirPath = GetDirPath(sTrgtFilePath)
                    sTrgtFileName = GetFileName(sTrgtFilePath)
                    Set oTrgtDirFiles = WScript.CreateObject("Shell.Application").Namespace( sTrgtDirPath & "\" )
                    Set oTrgtFile = oTrgtDirFiles.ParseName( sTrgtFileName )
                    sTrgtModDate = oTrgtFile.ModifyDate
                End If
                
                'タグ更新
                objLogFile.WriteLine objTrack.Location & chr(9) & objTrack.Grouping & chr(9) & objTrack.Composer
                
                '更新日を変更するために「ジャンル」を書き換え＋書き戻し
                'objTrack.VolumeAdjustment = 0
                Dim sTestGenre
                sTestGenre = objTrack.Genre
                objTrack.Genre = "Test"
                'objTrack.Genre = sTestGenre
                
                objTrack.Grouping = ""
                objTrack.Composer = ""
                
                '更新日書き戻し
                If UPDATE_MODIFIED_DATE = True Then
                    '更新日は更新されたまま
                Else
                    oTrgtFile.ModifyDate = CDate( sTrgtModDate ) '更新日を書き戻し
                End If
                lHitFileNum = lHitFileNum + 1
            End If
        Else
            'Do Nothing
        End If
        lTrgtFileListIdx = lTrgtFileListIdx + 1
    Next
    Set objPlayList = Nothing
    Set oTrgtFile = Nothing
    Set oTrgtDirFiles = Nothing
    
    objLogFile.WriteLine "ファイル数：" & lHitFileNum
    objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime
    objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime
Else
    objLogFile.WriteLine "経過時間（本処理のみ） : " & oStpWtch.IntervalTime
    objLogFile.WriteLine "経過時間（総時間）     : " & oStpWtch.ElapsedTime
End If ' ★Debug★

' ******************************************
' * 終了処理                               *
' ******************************************
Call Finish
WScript.CreateObject("WScript.Shell").Run sLogFilePath, 1, True
MsgBox "プログラムが正常に終了しました。"

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

'==========================================================
'= インクルード関数
'==========================================================
Private Function Include( _
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

