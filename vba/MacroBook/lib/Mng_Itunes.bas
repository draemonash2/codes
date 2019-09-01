Attribute VB_Name = "Mng_Itunes"
Option Explicit

' itunes libary v1.11

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

' ******************************************************************
' * グローバル関数
' ******************************************************************
' ==================================================================
' = 概要    iTunes プレイリストをバックアップする
' = 引数    なし
' = 戻値                String          バックアップディレクトリパス
' = 覚書    なし
' = 依存    Mng_FileSys.bas/CreateDirectry()
' = 所属    Mng_Itunes.bas
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
' = 依存    なし
' = 所属    Mng_Itunes.bas
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
                Case "歌詞": sTagValue = .Lyrics
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
' = 依存    なし
' = 所属    Mng_Itunes.bas
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
                Case "歌詞": .Lyrics = sTagValue
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
' = 概要    ファイルパスからトラック情報を取得する
' = 引数    sInTrgtPath         String    [in]  取得したいトラックのトラック名
' = 引数    vOutTrackInfo       Variant   [out] トラック情報オブジェクト
' = 引数    sOutErrorDetail     String    [out] エラー詳細情報
' = 引数    lInFileInfoTagIndex Long      [in]  ファイル取得情報種別番号
' = 戻値                        Boolean         取得結果
' = 覚書    ・以下のいずれかを満たす場合、取得結果 False を返却する。
' =           - ファイルパスが空
' =               OutErrorDetail ： "File path is empty!"
' =           - ファイル自体が存在しない場合
' =               OutErrorDetail ： "File is not exist at file system!"
' =           - ファイルが iTunes ライブラリ内に存在しない
' =               sOutErrorDetail ： "File is not exist at itunes playlist!"
' =         ・本関数使用時は FileSystem.bas をインポートしておくこと。
' =         ・本関数初回呼出前には ItunesInit() を実行しておくこと。
' =         ・lFileInfoTagIndex の指定がない場合は自動的に関数内で取得し､
' =           指定した場合は取得処理を省略する｡そのため､lFileInfoTagIndex が
' =           あらかじめわかっている場合､引数を省略することで処理を高速化できる｡
' =           使い方としては、本関数初回呼び出し前に FileSystem.bas の
' =           GetFileDetailInfoIndex() にて lInFileInfoTagIndex を取得して
' =           おいてから本関数を呼び出すとよい｡
' = 依存    Mng_FileInfo.bas/GetFileDetailInfoIndex()
' =         Mng_FileInfo.bas/GetFileDetailInfo()
' =         Mng_Itunes.bas/SearchTrack()
' = 所属    Mng_Itunes.bas
' ==================================================================
Public Function GetTrackInfo( _
    ByVal sInTrgtPath As String, _
    ByRef vOutTrackInfo As Variant, _
    ByRef sOutErrorDetail As String, _
    Optional ByVal lInFileInfoTagIndex As Long = 9999 _
) As Boolean
    Const FILE_DETAIL_INFO_TRACK_NAME_TITLE As String = "タイトル"
    
    sOutErrorDetail = ""
    Set vOutTrackInfo = Nothing
    
    Dim bIsError As Boolean
    bIsError = False
    
    '引数チェック
    If bIsError = True Then
        'Do Nothing
    Else
        If sInTrgtPath = "" Then
            bIsError = True
            sOutErrorDetail = "File path is empty!"
        Else
            'Do Nothing
        End If
    End If
    
    '「FILE_DETAIL_INFO_TRACK_NAME_TITLE」のインデックス取得
    If bIsError = True Then
        'Do Nothing
    Else
        If lInFileInfoTagIndex = 9999 Then
            'lInFileInfoTagIndex が指定されていない場合、lFileInfoTagIndex を取得する
            Dim bRet As Boolean
            bRet = GetFileDetailInfoIndex(FILE_DETAIL_INFO_TRACK_NAME_TITLE, lInFileInfoTagIndex)
            If bRet = True Then
                'Do Nothing
            Else
                Debug.Assert 0
            End If
        Else
            'lInFileInfoTagIndex が指定されている場合、指定された lInFileInfoTagIndex にて以降の処理を行う
        End If
    End If
    
    'トラック名取得
    If bIsError = True Then
        'Do Nothing
    Else
        Dim vFileInfoValue As Variant
        Dim vFileInfoTitle As Variant
        Dim sErrorDetail As String
        Dim sTrackName As String
        bRet = GetFileDetailInfo(sInTrgtPath, lInFileInfoTagIndex, vFileInfoValue, vFileInfoTitle, sErrorDetail)
        If bRet = True Then
            If vFileInfoTitle = FILE_DETAIL_INFO_TRACK_NAME_TITLE Then
                sTrackName = CStr(vFileInfoValue)
            Else
                Debug.Assert 0
            End If
        Else
            If sErrorDetail = "File is not exist!" Then
                bIsError = True
                sOutErrorDetail = "File is not exist at file system!"
            ElseIf sErrorDetail = "Get info type error!" Then
                Debug.Assert 0
            Else
                Debug.Assert 0
            End If
        End If
    End If
    
    'トラック情報取得
    If bIsError = True Then
        'Do Nothing
    Else
        bRet = SearchTrack(sTrackName, sInTrgtPath, vOutTrackInfo)
        If bRet = True Then
            'Do Nothing
        Else
            bIsError = True
            sOutErrorDetail = "File is not exist at itunes playlist!"
        End If
    End If
    
    If bIsError = True Then
        GetTrackInfo = False
    Else
        GetTrackInfo = True
    End If
End Function
    Private Sub Test_GetTrackInfo()
        Call ItunesInit
        
        Dim sTrgtPath As String
        Dim bRet As Boolean
        Dim sErrorDetail As String
        Dim vTrackInfo As Variant
        
        sTrgtPath = "Z:\300_Musics\999_Other\test\test album\01 test track 1.mp3"
        bRet = GetTrackInfo(sTrgtPath, vTrackInfo, sErrorDetail)
        If vTrackInfo Is Nothing Then
            Debug.Print "[" & bRet & " : " & sErrorDetail & "] "
        Else
            Debug.Print "[" & bRet & " : " & sErrorDetail & "] " & vTrackInfo.TrackNumber
        End If
        
        sTrgtPath = "Z:\300_Musics\999_Other\test\test album\15 test track 15.mp3"
        bRet = GetTrackInfo(sTrgtPath, vTrackInfo, sErrorDetail)
        If vTrackInfo Is Nothing Then
            Debug.Print "[" & bRet & " : " & sErrorDetail & "] "
        Else
            Debug.Print "[" & bRet & " : " & sErrorDetail & "] " & vTrackInfo.TrackNumber
        End If
        
        sTrgtPath = "Z:\300_Musics\999_Other\test\test album\16 test track 16.mp3"
        bRet = GetTrackInfo(sTrgtPath, vTrackInfo, sErrorDetail)
        If vTrackInfo Is Nothing Then
            Debug.Print "[" & bRet & " : " & sErrorDetail & "] "
        Else
            Debug.Print "[" & bRet & " : " & sErrorDetail & "] " & vTrackInfo.TrackNumber
        End If
        
        sTrgtPath = ""
        bRet = GetTrackInfo(sTrgtPath, vTrackInfo, sErrorDetail)
        If vTrackInfo Is Nothing Then
            Debug.Print "[" & bRet & " : " & sErrorDetail & "] "
        Else
            Debug.Print "[" & bRet & " : " & sErrorDetail & "] " & vTrackInfo.TrackNumber
        End If
        
        Call ItunesTerminate
    End Sub

' ******************************************************************
' * ローカル関数
' ******************************************************************
' ==================================================================
' = 概要    トラック情報オブジェクト取得
' = 引数    sTrackName    String    [in]    取得したいトラックのトラック名
' = 引数    sFilePath     String    [in]    取得したいトラックのファイルパス
' = 引数    vTrack        Variant   [out]   トラック情報オブジェクト
' = 戻値                  Boolean           取得結果
' = 覚書    取得したいトラックが見つからなかった場合、取得結果 False を返却する。
' = 依存    なし
' = 所属    Mng_Itunes.bas
' ==================================================================
Private Function SearchTrack( _
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

'Private Sub test()
'    Call ItunesInit
'    
'    Dim sTrgtPath As String
'    Dim bRet As Boolean
'    Dim sErrorDetail As String
'    Dim vTrackInfo As Variant
'    
'    sTrgtPath = "Z:\300_Musics\999_Other\test\test album\01 test track 1.mp3"
'    'sTrgtPath = "Z:\300_Musics\290_Reggae@Riddim\Jim Screechie\Riddim Mix (Jim Screechie Riddim).mp3"
'    bRet = GetTrackInfo(sTrgtPath, vTrackInfo, sErrorDetail)
'    'vTrackInfo.Lyrics = "aaa"
'    Debug.Print vTrackInfo.Lyrics
'    
'    Call ItunesTerminate
'End Sub

