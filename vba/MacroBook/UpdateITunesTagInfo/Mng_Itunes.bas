Attribute VB_Name = "Mng_Itunes"
Option Explicit

' itunes libary v1.0

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

' ==================================================================
' = �T�v    iTunes �v���C���X�g���o�b�N�A�b�v����
' = ����    �Ȃ�
' = �ߒl                String          �o�b�N�A�b�v�f�B���N�g���p�X
' = �o��    �Ȃ�
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
' = �T�v    �g���b�N���I�u�W�F�N�g�擾
' = ����    sTrackName    String    [in]    �擾�������g���b�N�̃g���b�N��
' = ����    sFilePath     String    [in]    �擾�������g���b�N�̃t�@�C���p�X
' = ����    vTrack      Variant   [out]   �g���b�N���I�u�W�F�N�g
' = �ߒl                  Boolean           �擾����
' = �o��    �擾�������g���b�N��������Ȃ������ꍇ�A�擾���� False ��ԋp����B
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
            MsgBox "�g���b�N��������܂���ł���"
        End If
        
        Call ItunesTerminate
    End Function

