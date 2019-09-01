Attribute VB_Name = "Mng_Itunes"
Option Explicit

' itunes libary v1.11

Public gvItunes As Variant
Public gvPlayList As Variant

Public Function ItunesInit()
    Set gvItunes = CreateObject("iTunes.Application")
    Set gvPlayList = gvItunes.Sources.Item(1).Playlists.ItemByName("�~���[�W�b�N")
End Function

Public Function ItunesTerminate()
    Set gvItunes = Nothing
    Set gvPlayList = Nothing
End Function

' ******************************************************************
' * �O���[�o���֐�
' ******************************************************************
' ==================================================================
' = �T�v    iTunes �v���C���X�g���o�b�N�A�b�v����
' = ����    �Ȃ�
' = �ߒl                String          �o�b�N�A�b�v�f�B���N�g���p�X
' = �o��    �Ȃ�
' = �ˑ�    Mng_FileSys.bas/CreateDirectry()
' = ����    Mng_Itunes.bas
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
' = �T�v    �^�O���擾����
' = ����    vTrack        Variant   [in]    �g���b�N���I�u�W�F�N�g
' = ����    sTagTitle     String    [in]    �擾�������^�O�̃^�O��
' = ����    sTagValue     String    [out]   �擾�����^�O�̒l
' = �ߒl                  Boolean           �擾����
' = �o��    �w�肵���^�O��������Ȃ������ꍇ�A�܂��́AobjTrack ��
' =         ��I�u�W�F�N�g�̏ꍇ�A�擾���� False ��ԋp����B
' = �ˑ�    �Ȃ�
' = ����    Mng_Itunes.bas
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
                Case "�A���o����": sTagValue = .Album
                Case "�A���o���A�[�e�B�X�g��": sTagValue = .AlbumArtist
                Case "�A�[�e�B�X�g��": sTagValue = .Artist
                Case "�g���b�N��": sTagValue = .Name
                Case "BPM": sTagValue = .BPM
                Case "�R�����g": sTagValue = .Comment
                Case "�R���s���[�V����": sTagValue = .Compilation
                Case "��Ȏ�": sTagValue = .Composer
                Case "�f�B�X�N��": sTagValue = .DiscCount
                Case "�f�B�X�N�ԍ�": sTagValue = .DiscNumber
                Case "�L��": sTagValue = .Enabled
                Case "�C�R���C�U": sTagValue = .EQ
                Case "�W������": sTagValue = .Genre
                Case "�O���[�v": sTagValue = .Grouping
                Case "�Đ���": sTagValue = .PlayedCount
                Case "�Đ���": sTagValue = .PlayedDate
                Case "�ǉ���": sTagValue = .DateAdded
                Case "�ύX��": sTagValue = .ModificationDate
                Case "���[�e�B���O": sTagValue = .Rating
                Case "�J�n����": sTagValue = .Start
                Case "�I������": sTagValue = .Finish
                Case "�g���b�N��": sTagValue = .TrackCount
                Case "�g���b�N�ԍ�": sTagValue = .TrackNumber
                Case "���ʒ���": sTagValue = .VolumeAdjustment
                Case "�N": sTagValue = .Year
                Case "�t�@�C���p�X": sTagValue = .Location
                Case "��ށi���l�j": sTagValue = .Kind
                Case "��ށi������j": sTagValue = .KindAsString
                Case "�r�b�g���[�g": sTagValue = .BitRate
                Case "�T���v�����[�g": sTagValue = .SampleRate
                Case "�T�C�Y": sTagValue = .Size
                Case "����(MIN)": sTagValue = .Time
                Case "����(SECOND)": sTagValue = .Duration
                Case "�Đ����C���f�b�N�X": sTagValue = .PlayOrderIndex
                Case "�C���f�b�N�X": sTagValue = .Index
                Case "�̎�": sTagValue = .Lyrics
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
            MsgBox "�G���["
            End
        End If
        
        Dim sTagTitle As String
        Dim sTagValue As String
        sTagTitle = "�A���o���A�[�e�B�X�g��"
        bRet = GetTagValue(vTrack, sTagTitle, sTagValue)
        
        Call ItunesTerminate
        
        MsgBox sTagValue
    End Sub

' ==================================================================
' = �T�v    �^�O���X�V����
' = ����    vTrack        Variant   [in]    �g���b�N���I�u�W�F�N�g
' = ����    sTagTitle     String    [in]    �X�V�������^�O�̃^�O��
' = ����    sTagValue     String    [in]    �X�V�������^�O�̒l
' = �ߒl                  Boolean           �擾����
' = �o��    �w�肵���^�O��������Ȃ������ꍇ�A�܂��́AobjTrack ��
' =         ��I�u�W�F�N�g�̏ꍇ�A�擾���� False ��ԋp����B
' = �ˑ�    �Ȃ�
' = ����    Mng_Itunes.bas
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
                Case "�A���o����": .Album = sTagValue
                Case "�A���o���A�[�e�B�X�g��": .AlbumArtist = sTagValue
                Case "�A�[�e�B�X�g��": .Artist = sTagValue
                Case "�g���b�N��": .Name = sTagValue
                Case "BPM": .BPM = sTagValue
                Case "�R�����g": .Comment = sTagValue
                Case "�R���s���[�V����": .Compilation = sTagValue
                Case "��Ȏ�": .Composer = sTagValue
                Case "�f�B�X�N��": .DiscCount = sTagValue
                Case "�f�B�X�N�ԍ�": .DiscNumber = sTagValue
                Case "�L��": .Enabled = sTagValue
                Case "�C�R���C�U": .EQ = sTagValue
                Case "�W������": .Genre = sTagValue
                Case "�O���[�v": .Grouping = sTagValue
                Case "�Đ���": .PlayedCount = sTagValue
                Case "�Đ���": .PlayedDate = sTagValue
                Case "���[�e�B���O": .Rating = sTagValue
                Case "�J�n����": .Start = sTagValue
                Case "�I������": .Finish = sTagValue
                Case "�g���b�N��": .TrackCount = sTagValue
                Case "�g���b�N�ԍ�": .TrackNumber = sTagValue
                Case "���ʒ���": .VolumeAdjustment = sTagValue
                Case "�N": .Year = sTagValue
                Case "�t�@�C���p�X": .Location = sTagValue
                Case "�̎�": .Lyrics = sTagValue
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
            MsgBox "�G���["
            End
        End If
        
        Dim sTagTitle As String
        Dim sTagValue As String
        sTagTitle = "�A���o���A�[�e�B�X�g��"
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
            MsgBox "�G���["
            End
        End If
        
        Dim sTagTitle As String
        Dim sTagValue As String
        sTagTitle = "�A���o���A�[�e�B�X�g��"
        sTagValue = "test albumartist"
        
        bRet = SetTagValue(vTrack, sTagTitle, sTagValue)
        MsgBox bRet
        
        Call ItunesTerminate
    End Sub

' ==================================================================
' = �T�v    �t�@�C���p�X����g���b�N�����擾����
' = ����    sInTrgtPath         String    [in]  �擾�������g���b�N�̃g���b�N��
' = ����    vOutTrackInfo       Variant   [out] �g���b�N���I�u�W�F�N�g
' = ����    sOutErrorDetail     String    [out] �G���[�ڍ׏��
' = ����    lInFileInfoTagIndex Long      [in]  �t�@�C���擾����ʔԍ�
' = �ߒl                        Boolean         �擾����
' = �o��    �E�ȉ��̂����ꂩ�𖞂����ꍇ�A�擾���� False ��ԋp����B
' =           - �t�@�C���p�X����
' =               OutErrorDetail �F "File path is empty!"
' =           - �t�@�C�����̂����݂��Ȃ��ꍇ
' =               OutErrorDetail �F "File is not exist at file system!"
' =           - �t�@�C���� iTunes ���C�u�������ɑ��݂��Ȃ�
' =               sOutErrorDetail �F "File is not exist at itunes playlist!"
' =         �E�{�֐��g�p���� FileSystem.bas ���C���|�[�g���Ă������ƁB
' =         �E�{�֐�����ďo�O�ɂ� ItunesInit() �����s���Ă������ƁB
' =         �ElFileInfoTagIndex �̎w�肪�Ȃ��ꍇ�͎����I�Ɋ֐����Ŏ擾���
' =           �w�肵���ꍇ�͎擾�������ȗ����顂��̂��ߤlFileInfoTagIndex ��
' =           ���炩���߂킩���Ă���ꍇ��������ȗ����邱�Ƃŏ������������ł���
' =           �g�����Ƃ��ẮA�{�֐�����Ăяo���O�� FileSystem.bas ��
' =           GetFileDetailInfoIndex() �ɂ� lInFileInfoTagIndex ���擾����
' =           �����Ă���{�֐����Ăяo���Ƃ悢�
' = �ˑ�    Mng_FileInfo.bas/GetFileDetailInfoIndex()
' =         Mng_FileInfo.bas/GetFileDetailInfo()
' =         Mng_Itunes.bas/SearchTrack()
' = ����    Mng_Itunes.bas
' ==================================================================
Public Function GetTrackInfo( _
    ByVal sInTrgtPath As String, _
    ByRef vOutTrackInfo As Variant, _
    ByRef sOutErrorDetail As String, _
    Optional ByVal lInFileInfoTagIndex As Long = 9999 _
) As Boolean
    Const FILE_DETAIL_INFO_TRACK_NAME_TITLE As String = "�^�C�g��"
    
    sOutErrorDetail = ""
    Set vOutTrackInfo = Nothing
    
    Dim bIsError As Boolean
    bIsError = False
    
    '�����`�F�b�N
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
    
    '�uFILE_DETAIL_INFO_TRACK_NAME_TITLE�v�̃C���f�b�N�X�擾
    If bIsError = True Then
        'Do Nothing
    Else
        If lInFileInfoTagIndex = 9999 Then
            'lInFileInfoTagIndex ���w�肳��Ă��Ȃ��ꍇ�AlFileInfoTagIndex ���擾����
            Dim bRet As Boolean
            bRet = GetFileDetailInfoIndex(FILE_DETAIL_INFO_TRACK_NAME_TITLE, lInFileInfoTagIndex)
            If bRet = True Then
                'Do Nothing
            Else
                Debug.Assert 0
            End If
        Else
            'lInFileInfoTagIndex ���w�肳��Ă���ꍇ�A�w�肳�ꂽ lInFileInfoTagIndex �ɂĈȍ~�̏������s��
        End If
    End If
    
    '�g���b�N���擾
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
    
    '�g���b�N���擾
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
' * ���[�J���֐�
' ******************************************************************
' ==================================================================
' = �T�v    �g���b�N���I�u�W�F�N�g�擾
' = ����    sTrackName    String    [in]    �擾�������g���b�N�̃g���b�N��
' = ����    sFilePath     String    [in]    �擾�������g���b�N�̃t�@�C���p�X
' = ����    vTrack        Variant   [out]   �g���b�N���I�u�W�F�N�g
' = �ߒl                  Boolean           �擾����
' = �o��    �擾�������g���b�N��������Ȃ������ꍇ�A�擾���� False ��ԋp����B
' = �ˑ�    �Ȃ�
' = ����    Mng_Itunes.bas
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
            MsgBox "�g���b�N��������܂���ł���"
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

