Attribute VB_Name = "Mng_Itunes"
Option Explicit

' itunes libary v1.0

Public gvItunes As Variant
Public gvPlayList As Variant

Public Function ItunesInit()
    Set gvItunes = CreateObject("iTunes.Application")
    Set gvPlayList = gvItunes.Sources.Item(1).Playlists.ItemByName("ミュージック")
End Function

Public Function ItunesTerminate()
    Set gvItunes = Nothing
    Set gvPlayList = Nothing
End Function

' ==================================================================
' = 概要    iTunes プレイリストをバックアップする
' = 引数    なし
' = 戻値                String          バックアップディレクトリパス
' = 覚書    なし
' ==================================================================
Public Function BackUpItunesPlaylist( _
    Optional ByVal sBackupDirName As String _
) As String
    Dim sDate As String
    sDate = Format(Now, "yyyymmdd_hhmmss")
    
    Dim sBackupOrgDirPath As String
    Dim sBackupDstDirPath As String
    Dim sLogFilePath As String
    sBackupOrgDirPath = Replace(gvItunes.LibraryXMLPath, "\iTunes Music Library.xml", "")
    If sBackupDirName = "" Then
        sBackupDstDirPath = sBackupOrgDirPath & "\iTunes Library Backup\" & sDate
    Else
        sBackupDstDirPath = sBackupOrgDirPath & "\" & sBackupDirName & "\" & sDate
    End If
    
    Call CreateDirectry(sBackupDstDirPath)
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFile (sBackupOrgDirPath & "\iTunes Library Extras.itdb"), (sBackupDstDirPath & "\iTunes Library Extras.itdb")
    objFSO.CopyFile (sBackupOrgDirPath & "\iTunes Library Genius.itdb"), (sBackupDstDirPath & "\iTunes Library Genius.itdb")
    objFSO.CopyFile (sBackupOrgDirPath & "\iTunes Library.itl"), (sBackupDstDirPath & "\iTunes Library.itl")
    objFSO.CopyFile (sBackupOrgDirPath & "\iTunes Music Library.xml"), (sBackupDstDirPath & "\iTunes Music Library.xml")
    
    BackUpItunesPlaylist = sBackupDstDirPath
End Function
    Private Sub Test_BackUpItunesPlaylist()
        Call ItunesInit
        Debug.Print BackUpItunesPlaylist()
        Debug.Print BackUpItunesPlaylist("backup")
        Call ItunesTerminate
    End Sub

' ==================================================================
' = 概要    タグを取得する
' = 引数    vTrack        Variant   [in]    トラック情報オブジェクト
' = 引数    sTagTitle     String    [in]    取得したいタグのタグ名
' = 引数    sTagValue     String    [out]   取得したタグの値
' = 戻値                  Boolean           取得結果
' = 覚書    指定したタグが見つからなかった場合、または、objTrack が
' =         空オブジェクトの場合、取得結果 False を返却する。
' ==================================================================
Public Function GetTagValue( _
    ByRef vTrack As Variant, _
    ByVal sTagTitle As String, _
    ByRef sTagValue As String _
) As Boolean
    Dim bRet As Boolean
    bRet = True
    
    If bRet = True Then
        If vTrack Is Nothing Then
            bRet = False
        Else
            'Do Nothing
        End If
    Else
        'Do Nothing
    End If
    
    If bRet = True Then
        With vTrack
            Select Case sTagTitle
                Case "アルバム名": sTagValue = .Album
                Case "アルバムアーティスト名": sTagValue = .AlbumArtist
                Case "アーティスト名": sTagValue = .Artist
                Case "トラック名": sTagValue = .Name
                Case "BPM": sTagValue = .BPM
                Case "コメント": sTagValue = .Comment
                Case "コンピレーション": sTagValue = .Compilation
                Case "作曲者": sTagValue = .Composer
                Case "ディスク数": sTagValue = .DiscCount
                Case "ディスク番号": sTagValue = .DiscNumber
                Case "有効": sTagValue = .Enabled
                Case "イコライザ": sTagValue = .EQ
                Case "ジャンル": sTagValue = .Genre
                Case "グループ": sTagValue = .Grouping
                Case "再生回数": sTagValue = .PlayedCount
                Case "再生日": sTagValue = .PlayedDate
                Case "追加日": sTagValue = .DateAdded
                Case "変更日": sTagValue = .ModificationDate
                Case "レーティング": sTagValue = .Rating
                Case "開始時間": sTagValue = .Start
                Case "終了時間": sTagValue = .Finish
                Case "トラック数": sTagValue = .TrackCount
                Case "トラック番号": sTagValue = .TrackNumber
                Case "音量調整": sTagValue = .VolumeAdjustment
                Case "年": sTagValue = .Year
                Case "ファイルパス": sTagValue = .Location
                Case "種類（数値）": sTagValue = .Kind
                Case "種類（文字列）": sTagValue = .KindAsString
                Case "ビットレート": sTagValue = .BitRate
                Case "サンプルレート": sTagValue = .SampleRate
                Case "サイズ": sTagValue = .Size
                Case "時間(MIN)": sTagValue = .Time
                Case "時間(SECOND)": sTagValue = .Duration
                Case "再生順インデックス": sTagValue = .PlayOrderIndex
                Case "インデックス": sTagValue = .Index
                Case Else: sTagValue = "": bRet = False
            End Select
        End With
    Else
        'Do Nothings
    End If
    GetTagValue = bRet
End Function
    Private Sub Test_GetTagValue()
        Call ItunesInit
        
        Dim sTrackName As String
        Dim sFilePath As String
        sTrackName = "test track"
        sFilePath = "Z:\300_Musics\999_Other\test\test mp3 file.mp3"
        
        Dim vTrack As Variant
        Dim bRet As Boolean
        bRet = SearchTrack(sTrackName, sFilePath, vTrack)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox "エラー"
            End
        End If
        
        Dim sTagTitle As String
        Dim sTagValue As String
        sTagTitle = "アルバムアーティスト名"
        bRet = GetTagValue(vTrack, sTagTitle, sTagValue)
        
        Call ItunesTerminate
        
        MsgBox sTagValue
    End Sub

' ==================================================================
' = 概要    タグを更新する
' = 引数    vTrack        Variant   [in]    トラック情報オブジェクト
' = 引数    sTagTitle     String    [in]    更新したいタグのタグ名
' = 引数    sTagValue     String    [in]    更新したいタグの値
' = 戻値                  Boolean           取得結果
' = 覚書    指定したタグが見つからなかった場合、または、objTrack が
' =         空オブジェクトの場合、取得結果 False を返却する。
' ==================================================================
Public Function SetTagValue( _
    ByVal vTrack As Variant, _
    ByVal sTagTitle As String, _
    ByVal sTagValue As String _
) As Boolean
    Dim bRet As Boolean
    bRet = True
    
    If bRet = True Then
        If vTrack Is Nothing Then
            bRet = False
        Else
            'Do Nothing
        End If
    Else
        'Do Nothing
    End If
    
    If bRet = True Then
        With vTrack
            Select Case sTagTitle
                Case "アルバム名": .Album = sTagValue
                Case "アルバムアーティスト名": .AlbumArtist = sTagValue
                Case "アーティスト名": .Artist = sTagValue
                Case "トラック名": .Name = sTagValue
                Case "BPM": .BPM = sTagValue
                Case "コメント": .Comment = sTagValue
                Case "コンピレーション": .Compilation = sTagValue
                Case "作曲者": .Composer = sTagValue
                Case "ディスク数": .DiscCount = sTagValue
                Case "ディスク番号": .DiscNumber = sTagValue
                Case "有効": .Enabled = sTagValue
                Case "イコライザ": .EQ = sTagValue
                Case "ジャンル": .Genre = sTagValue
                Case "グループ": .Grouping = sTagValue
                Case "再生回数": .PlayedCount = sTagValue
                Case "再生日": .PlayedDate = sTagValue
                Case "レーティング": .Rating = sTagValue
                Case "開始時間": .Start = sTagValue
                Case "終了時間": .Finish = sTagValue
                Case "トラック数": .TrackCount = sTagValue
                Case "トラック番号": .TrackNumber = sTagValue
                Case "音量調整": .VolumeAdjustment = sTagValue
                Case "年": .Year = sTagValue
                Case "ファイルパス": .Location = sTagValue
                Case Else: bRet = False
            End Select
        End With
    Else
        'Do Nothing
    End If
    SetTagValue = bRet
End Function
    Private Sub Test_SetTagValue()
        Call ItunesInit
        
        Dim sTrackName As String
        Dim sFilePath As String
        sTrackName = "test track"
        sFilePath = "Z:\300_Musics\999_Other\test\test mp3 file.mp3"
        
        Dim vTrack As Variant
        Dim bRet As Boolean
        bRet = SearchTrack(sTrackName, sFilePath, vTrack)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox "エラー"
            End
        End If
        
        Dim sTagTitle As String
        Dim sTagValue As String
        sTagTitle = "アルバムアーティスト名"
        sTagValue = "endo"
        
        bRet = SetTagValue(vTrack, sTagTitle, sTagValue)
        MsgBox bRet
        
        Call ItunesTerminate
    End Sub
    Private Sub Test_SetTagValue02()
        Call ItunesInit
        
        Dim sTrackName As String
        Dim sFilePath As String
        sTrackName = "test track"
        sFilePath = "Z:\300_Musics\999_Other\test\test mp3 file.mp3"
        
        Dim vTrack As Variant
        Dim bRet As Boolean
        bRet = SearchTrack(sTrackName, sFilePath, vTrack)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox "エラー"
            End
        End If
        
        Dim sTagTitle As String
        Dim sTagValue As String
        sTagTitle = "アルバムアーティスト名"
        sTagValue = "test albumartist"
        
        bRet = SetTagValue(vTrack, sTagTitle, sTagValue)
        MsgBox bRet
        
        Call ItunesTerminate
    End Sub

' ==================================================================
' = 概要    トラック情報オブジェクト取得
' = 引数    sTrackName    String    [in]    取得したいトラックのトラック名
' = 引数    sFilePath     String    [in]    取得したいトラックのファイルパス
' = 引数    vTrack      Variant   [out]   トラック情報オブジェクト
' = 戻値                  Boolean           取得結果
' = 覚書    取得したいトラックが見つからなかった場合、取得結果 False を返却する。
' ==================================================================
Public Function SearchTrack( _
    ByVal sTrackName As String, _
    ByVal sFilePath As String, _
    ByRef vTrack As Variant _
) As Boolean
    If gvPlayList Is Nothing Then
        SearchTrack = False
        Exit Function
    Else
        'Do Nothing
    End If
    
    Dim objSearchResult As Variant
    Set objSearchResult = gvPlayList.Search(sTrackName, 5)
    
    Set vTrack = Nothing
    
    Dim bRet As Boolean
    If objSearchResult Is Nothing Then
        bRet = False
    Else
        bRet = False
        Dim lHitIdx As Long
        For lHitIdx = 1 To objSearchResult.Count
            If LCase(objSearchResult.Item(lHitIdx).Location) = LCase(sFilePath) Then
                Set vTrack = objSearchResult.Item(lHitIdx)
                bRet = True
                Exit For
            Else
                'Do Nothing
            End If
        Next
    End If
    SearchTrack = bRet
End Function
    Private Function Test_SearchTrack()
        Call ItunesInit
        
        Dim sTrackName As String
        Dim sFilePath As String
        sTrackName = "test track"
        sFilePath = "Z:\300_Musics\999_Other\test\test mp3 file.mp3"
        
        Dim vTrack As Variant
        Dim bRet As Boolean
        bRet = SearchTrack(sTrackName, sFilePath, vTrack)
        If bRet = True Then
            MsgBox vTrack.AlbumArtist
        Else
            MsgBox "トラックが見つかりませんでした"
        End If
        
        Call ItunesTerminate
    End Function

