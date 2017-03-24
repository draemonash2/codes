Attribute VB_Name = "Mng_Itunes"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim objItunes As Object
Dim objPlayList As Object
Public gsOrgDirPath As String
Public gsDstDirPath As String
Public gsDstDirBasePath As String
Public gsLogFilePath As String

Const ITUNES_BACKUP_FOLDER_MAX As Long = 20

Public Function ItunesInit()
    Set objItunes = CreateObject("iTunes.Application")
    'Set objPlayList = objItunes.Sources.Item(1).Playlists.ItemByName("ミュージック")
    Set objPlayList = objItunes.LibrarySource.Playlists.ItemByName("ミュージック")
    'Set objPlayList = objItunes.LibraryPlaylist
    'Set objPlayList = objItunes.LibrarySource
    'Set objPlayList = objItunes.Sources

    Dim sDate
    sDate = Format(Now, "yyyymmdd_hhmmss")
    gsOrgDirPath = RemoveTailWord(objItunes.LibraryXMLPath, "\")
    gsDstDirBasePath = gsOrgDirPath & "\iTunes Library Backup"
    gsDstDirPath = gsDstDirBasePath & "\" & sDate
    gsLogFilePath = gsDstDirPath & "\" & Replace(ThisWorkbook.Name, ".xlsm", ".log")
End Function

Public Function BackUpItunesPlaylist()
    '*** フォルダをバックアップ ***
    MkDir gsDstDirPath
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFile (gsOrgDirPath & "\iTunes Library Extras.itdb"), (gsDstDirPath & "\iTunes Library Extras.itdb")
    objFSO.CopyFile (gsOrgDirPath & "\iTunes Library Genius.itdb"), (gsDstDirPath & "\iTunes Library Genius.itdb")
    objFSO.CopyFile (gsOrgDirPath & "\iTunes Library.itl"), (gsDstDirPath & "\iTunes Library.itl")
    objFSO.CopyFile (gsOrgDirPath & "\iTunes Music Library.xml"), (gsDstDirPath & "\iTunes Music Library.xml")
    
    '*** 古いバックアップフォルダを削除 ***
    Dim asDirList As Variant
    Call GetFileList2(gsDstDirBasePath, asDirList, 2)
    
    'フォルダ削除
    If UBound(asDirList) >= ITUNES_BACKUP_FOLDER_MAX Then
        Dim lDelFolderMax
        lDelFolderMax = UBound(asDirList) - ITUNES_BACKUP_FOLDER_MAX
        Dim lDelDirIdx
        For lDelDirIdx = LBound(asDirList) To lDelFolderMax
            'バックアップフォルダ名は「YYYYMMDD_HHMMSS」で統一されているため、
            'asDirList() は自然と日時順に並ぶ。（要素番号が大きくなるほど新しい）
            'そのため、要素番号の小さい順からフォルダを削除する。
            objFSO.DeleteFolder asDirList(lDelDirIdx), True
        Next
    Else
        'Do Nothing
    End If
End Function

Public Function ReplaceItunesLibLocation()
    Open gsLogFilePath For Output As #1
    
    On Error Resume Next
    
    Print #1, "*** replace words ***"
    Print #1, "[source]" & Chr(9) & "[destination]"
    Dim sSrcLoc As String
    Dim sDstLoc As String
    Dim lRepInfoIdx As Long
    For lRepInfoIdx = 0 To UBound(gtBasicInfo.atReplaceInfo)
        sSrcLoc = gtBasicInfo.atReplaceInfo(lRepInfoIdx).sRepKeywordSrc
        sDstLoc = gtBasicInfo.atReplaceInfo(lRepInfoIdx).sRepKeywordDst
        Print #1, sSrcLoc & Chr(9) & sDstLoc
    Next lRepInfoIdx
    
    Print #1, ""
    Print #1, "*** replace results ***"
    Print #1, "  Replaced ) replace finished."
    Print #1, "  NotExist ) string matched, but file doesn't exist."
    Print #1, "  UnMatch  ) string unmatched."
    Print #1, "  NotMpeg  ) not local mp3 file."
    Dim sPathOrg As String
    Dim objTrack As Object
    Dim bIsMatch As Boolean
    Dim bIsFileExist As Boolean
    Dim lTracksNum As Long
    Dim lTracksIdx As Long
    lTracksNum = objPlayList.Tracks.Count
    lTracksIdx = 0
    For Each objTrack In objPlayList.Tracks
        If objTrack.KindAsString = "MPEG オーディオファイル" Or objTrack.KindAsString = "MPEGオーディオファイル" Then
            sPathOrg = objTrack.Location
            bIsMatch = False
            bIsFileExist = False
            For lRepInfoIdx = 0 To UBound(gtBasicInfo.atReplaceInfo)
                sSrcLoc = gtBasicInfo.atReplaceInfo(lRepInfoIdx).sRepKeywordSrc
                sDstLoc = gtBasicInfo.atReplaceInfo(lRepInfoIdx).sRepKeywordDst
                If InStr(sPathOrg, sSrcLoc) > 0 Then
                    objTrack.Location = Replace(sPathOrg, sSrcLoc, sDstLoc) 'sDstLocが存在しない場合、エラーが発生する。
                    If Err.Number = 0 Then
                        bIsMatch = True
                        bIsFileExist = True
                    Else
                        bIsMatch = True
                        Err.Clear
                    End If
                    Exit For
                Else
                    'Do Nothing
                End If
            Next lRepInfoIdx
            If bIsMatch = True Then
                If bIsFileExist = True Then
                    Print #1, "[Replaced] " & sPathOrg
                Else
                    Print #1, "[NotExist] " & sPathOrg
                End If
            Else
                Print #1, "[UnMatch ] " & sPathOrg
            End If
        Else
            Print #1, "[NotMpeg ] " & objTrack.Name & Chr(9) & objTrack.KindAsString
        End If
        
        'プログレスバー更新
        goPrgrsBar.Update (lTracksIdx / lTracksNum)
        If goPrgrsBar.IsCanceled = True Then
            Exit For
        Else
            'Do Nothing
        End If
        lTracksIdx = lTracksIdx + 1
    Next
    On Error GoTo 0
    
    Print #1, "TrackNum : " & objPlayList.Tracks.Count
    
    Close #1

End Function

Public Function OutputItunesLibLocation()
    gsLogFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\OutputLibPaths.log"
    Open gsLogFilePath For Output As #1
    
    On Error Resume Next
    
    Dim sPathOrg As String
    Dim objTrack As Object
    Dim bIsMatch As Boolean
    Dim bIsFileExist As Boolean
    Dim lTracksNum As Long
    Dim lTracksIdx As Long
    lTracksNum = objPlayList.Tracks.Count
    lTracksIdx = 0
    For Each objTrack In objPlayList.Tracks
        If objTrack.KindAsString = "MPEG オーディオファイル" Then
            Print #1, objTrack.Location
        Else
        End If
        
        'プログレスバー更新
        goPrgrsBar.Update (lTracksIdx / lTracksNum)
        If goPrgrsBar.IsCanceled = True Then
            Exit For
        Else
            'Do Nothing
        End If
        lTracksIdx = lTracksIdx + 1
    Next
    On Error GoTo 0
    
    Print #1, "TrackNum : " & objPlayList.Tracks.Count
    
    Close #1
End Function

Public Function ItunesTerminate()
    Set objItunes = Nothing
    Set objPlayList = Nothing
End Function

