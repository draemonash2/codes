Attribute VB_Name = "FileInfo"
Option Explicit

' file info libary v1.1

' CreateFile �֐�
Private Declare Function CreateFile Lib "KERNEL32.DLL" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

' CloseHandle �֐�
Private Declare Function CloseHandle Lib "KERNEL32.DLL" ( _
    ByVal hObject As Long _
) As Long

' LocalFileTimeToFileTime �֐�
Private Declare Function LocalFileTimeToFileTime Lib "KERNEL32.DLL" ( _
    ByRef lpLocalFileTime As FileTime, _
    ByRef lpFileTime As FileTime _
) As Long

' SystemTimeToFileTime �֐�
Private Declare Function SystemTimeToFileTime Lib "KERNEL32.DLL" ( _
    ByRef lpSystemTime As SystemTime, _
    ByRef lpFileTime As FileTime _
) As Long

' SetFileTime �֐�
Private Declare Function SetFileTime Lib "KERNEL32.DLL" ( _
    ByVal cFile As Long, _
    ByRef lpCreationTime As FileTime, _
    ByRef lpLastAccessTime As FileTime, _
    ByRef lpLastWriteTime As FileTime _
) As Long

' SystemTime �\����
Private Type SystemTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type

' FileTime �\����
Private Type FileTime
    LowDateTime As Long
    HighDateTime As Long
End Type

' �萔�̒�`
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const OPEN_EXISTING As Long = 3

' FileTime ���擾����
Private Function GetFileTime(ByVal dtSetting As Date) As FileTime
    Dim tSystemTime As SystemTime
    
    With tSystemTime
        .Year = Year(dtSetting)
        .Month = Month(dtSetting)
        .DayOfWeek = Weekday(dtSetting)
        .Day = Day(dtSetting)
        .Hour = Hour(dtSetting)
        .Minute = Minute(dtSetting)
        .Second = Second(dtSetting)
    End With
    
    Dim tLocalTime As FileTime
    Call SystemTimeToFileTime(tSystemTime, tLocalTime)
    
    Dim tFileTime As FileTime
    Call LocalFileTimeToFileTime(tLocalTime, tFileTime)
    
    GetFileTime = tFileTime
End Function

' �t�@�C���̃n���h�����擾����
Private Function GetFileHandle(ByVal stFilePath As String) As Long
    GetFileHandle = CreateFile( _
        stFilePath, GENERIC_READ Or GENERIC_WRITE, _
        FILE_SHARE_READ, 0, OPEN_EXISTING, _
        FILE_ATTRIBUTE_NORMAL, 0 _
    )
End Function

' -------------------------------------------------------------------------------
' �쐬�������w�肵�����t�Ǝ��Ԃɐݒ肵�܂��B
'
' @Param stFilePath �쐬������ݒ肷��t�@�C���܂ł̃p�X�B
' @Param stCreateTime �쐬�����ɐݒ肷����t�Ǝ��ԁB
' -------------------------------------------------------------------------------
Public Function SetCreationTime( _
    ByVal stFilePath As String, _
    ByVal stCreateTime As String _
)
    ' FileTime ���擾����
    Dim tFileTime As FileTime
    tFileTime = GetFileTime(CDate(stCreateTime))
    
    ' �t�@�C���̃n���h�����擾����
    Dim cFileHandle As Long
    cFileHandle = GetFileHandle(stFilePath)
    
    ' �t�@�C���̃n���h�����擾�ł����ꍇ�̂݁u�쐬�����v���X�V����
    If cFileHandle >= 0 Then
        Dim tNullable As FileTime
        
        Call SetFileTime(cFileHandle, tFileTime, tNullable, tNullable)
        Call CloseHandle(cFileHandle)
    End If
End Function

' -------------------------------------------------------------------------------
' �X�V�������w�肵�����t�Ǝ��Ԃɐݒ肵�܂��B
'
' @Param stFilePath �X�V������ݒ肷��t�@�C���܂ł̃p�X�B
' @Param stUpdateTime �X�V�����ɐݒ肷����t�Ǝ��ԁB
' -------------------------------------------------------------------------------
Public Function SetLastWriteTime( _
    ByVal stFilePath As String, _
    ByVal stUpdateTime As String _
)
    ' FileTime ���擾����
    Dim tFileTime As FileTime
    tFileTime = GetFileTime(CDate(stUpdateTime))
    
    ' �t�@�C���̃n���h�����擾����
    Dim cFileHandle As Long
    cFileHandle = GetFileHandle(stFilePath)
    
    ' �t�@�C���̃n���h�����擾�ł����ꍇ�̂݁u�X�V�����v���X�V����
    If cFileHandle >= 0 Then
        Dim tNullable As FileTime
        
        Call SetFileTime(cFileHandle, tNullable, tNullable, tFileTime)
        Call CloseHandle(cFileHandle)
    End If
End Function

' -------------------------------------------------------------------------------
' �A�N�Z�X�������w�肵�����t�Ǝ��Ԃɐݒ肵�܂��B
'
' @Param stFilePath �A�N�Z�X������ݒ肷��t�@�C���܂ł̃p�X�B
' @Param stAccessTime �A�N�Z�X�����ɐݒ肷����t�Ǝ��ԁB
' -------------------------------------------------------------------------------
Public Function SetLastAccessTime( _
    ByVal stFilePath As String, _
    ByVal stAccessTime As String _
)
    ' FileTime ���擾����
    Dim tFileTime As FileTime
    tFileTime = GetFileTime(CDate(stAccessTime))
    
    ' �t�@�C���̃n���h�����擾����
    Dim cFileHandle As Long
    cFileHandle = GetFileHandle(stFilePath)
    
    ' �t�@�C���̃n���h�����擾�ł����ꍇ�̂݁u�A�N�Z�X�����v���X�V����
    If cFileHandle >= 0 Then
        Dim tNullable As FileTime
        
        Call SetFileTime(cFileHandle, tNullable, tFileTime, tNullable)
        Call CloseHandle(cFileHandle)
    End If
End Function

Public Function GetCreationTime( _
    ByVal sFilePath As String _
) As String
    GetCreationTime = CreateObject("Scripting.FileSystemObject").GetFile(sFilePath).DateCreated
End Function

Public Function GetLastWriteTime( _
    ByVal sFilePath As String _
) As String
    GetLastWriteTime = CreateObject("Scripting.FileSystemObject").GetFile(sFilePath).DateLastModified
End Function

Public Function GetLastAccessTime( _
    ByVal sFilePath As String _
) As String
    GetLastAccessTime = CreateObject("Scripting.FileSystemObject").GetFile(sFilePath).DateLastAccessed
End Function

' ==================================================================
' = �T�v    �t�@�C�����擾
' = ����    sTrgtPath       String      [in]    �t�@�C���p�X
' = ����    lGetInfoType    Long        [in]    �擾����� (��1)
' = ����    vFileInfo       Variant     [out]   �t�@�C����� (��1)
' = �ߒl                    Boolean             �擾����
' = �o��    �ȉ��A�Q�ƁB
' =     (��1) �t�@�C�����
' =         [����]  [����]                  [�v���p�e�B��]      [�f�[�^�^]              [Get/Set]   [�o�͗�]
' =         1       �t�@�C����              Name                vbString    ������^    Get/Set     03 Ride Featuring Tony Matterhorn.MP3
' =         2       �t�@�C���T�C�Y          Size                vbLong      �������^    Get         4286923
' =         3       �t�@�C�����            Type                vbString    ������^    Get         MPEG layer 3
' =         4       �t�@�C���i�[��h���C�u  Drive               vbString    ������^    Get         Z:
' =         5       �t�@�C���p�X            Path                vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =         6       �e�t�H���_              ParentFolder        vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         7       MS-DOS�`���t�@�C����    ShortName           vbString    ������^    Get         03 Ride Featuring Tony Matterhorn.MP3
' =         8       MS-DOS�`���p�X          ShortPath           vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =         9       �쐬����                DateCreated         vbDate      ���t�^      Get         2015/08/19 0:54:45
' =         10      �A�N�Z�X����            DateLastAccessed    vbDate      ���t�^      Get         2016/10/14 6:00:30
' =         11      �X�V����                DateLastModified    vbDate      ���t�^      Get         2016/10/14 6:00:30
' =         12      ����                    Attributes          vbLong      �������^    (��2)       32
' =     (��2) ����
' =         [�l]                [����]                                      [������]    [Get/Set]
' =         1  �i0b00000001�j   �ǂݎ���p�t�@�C��                        ReadOnly    Get/Set
' =         2  �i0b00000010�j   �B���t�@�C��                                Hidden      Get/Set
' =         4  �i0b00000100�j   �V�X�e���E�t�@�C��                          System      Get/Set
' =         8  �i0b00001000�j   �f�B�X�N�h���C�u�E�{�����[���E���x��        Volume      Get
' =         16 �i0b00010000�j   �t�H���_�^�f�B���N�g��                      Directory   Get
' =         32 �i0b00100000�j   �O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1   Archive     Get/Set
' =         64 �i0b01000000�j   �����N�^�V���[�g�J�b�g                      Alias       Get
' =         128�i0b10000000�j   ���k�t�@�C��                                Compressed  Get
' ==================================================================
Public Function GetFileInfo( _
    ByVal sTrgtPath As String, _
    ByVal lGetInfoType As Long, _
    ByRef vFileInfo As Variant _
) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(sTrgtPath) Then
        'Do Nothing
    Else
        vFileInfo = ""
        GetFileInfo = False
        Exit Function
    End If
    
    Dim objFile As Object
    Set objFile = objFSO.GetFile(sTrgtPath)
    
    vFileInfo = ""
    GetFileInfo = True
    Select Case lGetInfoType
        Case 1:     vFileInfo = objFile.Name                '�t�@�C����
        Case 2:     vFileInfo = objFile.Size                '�t�@�C���T�C�Y
        Case 3:     vFileInfo = objFile.Type                '�t�@�C�����
        Case 4:     vFileInfo = objFile.Drive               '�t�@�C���i�[��h���C�u
        Case 5:     vFileInfo = objFile.Path                '�t�@�C���p�X
        Case 6:     vFileInfo = objFile.ParentFolder        '�e�t�H���_
        Case 7:     vFileInfo = objFile.ShortName           'MS-DOS�`���t�@�C����
        Case 8:     vFileInfo = objFile.ShortPath           'MS-DOS�`���p�X
        Case 9:     vFileInfo = objFile.DateCreated         '�쐬����
        Case 10:    vFileInfo = objFile.DateLastAccessed    '�A�N�Z�X����
        Case 11:    vFileInfo = objFile.DateLastModified    '�X�V����
        Case 12:    vFileInfo = objFile.Attributes          '����
        Case Else:  GetFileInfo = False
    End Select
End Function
    Private Sub Test_GetFileInfo()
        Dim sBuf As String
        Dim bRet As Boolean
        Dim vFileInfo As Variant
        sBuf = ""
        Dim sTrgtPath As String
        sTrgtPath = "C:\codes\vbs\lib\FileSystem.vbs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFileInfo(sTrgtPath, 1, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C�����F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 2, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���T�C�Y�F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 3, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C����ށF" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 4, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���i�[��h���C�u�F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 5, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���p�X�F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 6, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �e�t�H���_�F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 7, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  MS-DOS�`���t�@�C�����F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 8, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  MS-DOS�`���p�X�F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 9, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �쐬�����F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 10, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �A�N�Z�X�����F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 11, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �X�V�����F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 12, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �����F" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 13, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �F" & vFileInfo
        sTrgtPath = "C:\codes\vbs\lib\dummy.vbs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFileInfo(sTrgtPath, 1, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C�����F" & vFileInfo
        MsgBox sBuf
    End Sub

'�t�@�C�����́u�t�@�C�����v�u�����v���ݒ�\
'�������A�ȉ��̃��\�b�h�ɂĕύX�\�Ȃ��߁A���߂Ċ֐���`���Ȃ�
'  �t�@�C�����F objFSO.MoveFile
'  �����F objFSO.GetFile( "C:\codes\a.txt" ).Attributes
'Public Function SetFileInfo( _
'   ByVal sTrgtPath As String, _
'   ByVal lSetInfoType As Long, _
'   ByVal vFileInfo As Variant _
') As Boolean
'End Function

' ==================================================================
' = �T�v    �t�H���_���擾
' = ����    sTrgtPath       String      [in]    �t�H���_�p�X
' = ����    lGetInfoType    Long        [in]    �擾����� (��1)
' = ����    vFolderInfo     Variant     [out]   �t�H���_��� (��1)
' = �ߒl                    Boolean             �擾����
' = �o��    �ȉ��A�Q�ƁB
' =     (��1) �t�H���_���
' =         [����]  [����]                  [�v���p�e�B��]      [�f�[�^�^]              [Get/Set]   [�o�͗�]
' =         1       �t�H���_��              Name                vbString    ������^    Get/Set     Sacrifice
' =         2       �t�H���_�T�C�Y          Size                vbLong      �������^    Get         80613775
' =         3       �t�@�C�����            Type                vbString    ������^    Get         �t�@�C�� �t�H���_�[
' =         4       �t�@�C���i�[��h���C�u  Drive               vbString    ������^    Get         Z:
' =         5       �t�H���_�p�X            Path                vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         6       ���[�g �t�H���_         IsRootFolder        vbBoolean   �u�[���^    Get         False
' =         7       MS-DOS�`���t�@�C����    ShortName           vbString    ������^    Get         Sacrifice
' =         8       MS-DOS�`���p�X          ShortPath           vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         9       �쐬����                DateCreated         vbDate      ���t�^      Get         2015/08/19 0:54:44
' =         10      �A�N�Z�X����            DateLastAccessed    vbDate      ���t�^      Get         2015/08/19 0:54:44
' =         11      �X�V����                DateLastModified    vbDate      ���t�^      Get         2015/04/18 3:38:36
' =         12      ����                    Attributes          vbLong      �������^    (��2)       16
' =     (��2) ����
' =         [�l]                [����]                                      [������]    [Get/Set]
' =         1  �i0b00000001�j   �ǂݎ���p�t�@�C��                        ReadOnly    Get/Set
' =         2  �i0b00000010�j   �B���t�@�C��                                Hidden      Get/Set
' =         4  �i0b00000100�j   �V�X�e���E�t�@�C��                          System      Get/Set
' =         8  �i0b00001000�j   �f�B�X�N�h���C�u�E�{�����[���E���x��        Volume      Get
' =         16 �i0b00010000�j   �t�H���_�^�f�B���N�g��                      Directory   Get
' =         32 �i0b00100000�j   �O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1   Archive     Get/Set
' =         64 �i0b01000000�j   �����N�^�V���[�g�J�b�g                      Alias       Get
' =         128�i0b10000000�j   ���k�t�@�C��                                Compressed  Get
' ==================================================================
Public Function GetFolderInfo( _
    ByVal sTrgtPath As String, _
    ByVal lGetInfoType As Long, _
    ByRef vFolderInfo As Variant _
) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(sTrgtPath) Then
        'Do Nothing
    Else
        vFolderInfo = ""
        GetFolderInfo = False
        Exit Function
    End If
    
    Dim objFolder As Object
    Set objFolder = objFSO.GetFolder(sTrgtPath)
    
    vFolderInfo = ""
    GetFolderInfo = True
    Select Case lGetInfoType
        Case 1:     vFolderInfo = objFolder.Name                '�t�H���_��
        Case 2:     vFolderInfo = objFolder.Size                '�t�H���_�T�C�Y
        Case 3:     vFolderInfo = objFolder.Type                '�t�@�C�����
        Case 4:     vFolderInfo = objFolder.Drive               '�t�@�C���i�[��h���C�u
        Case 5:     vFolderInfo = objFolder.Path                '�t�H���_�p�X
        Case 6:     vFolderInfo = objFolder.IsRootFolder        '���[�g �t�H���_
        Case 7:     vFolderInfo = objFolder.ShortName           'MS-DOS�`���t�@�C����
        Case 8:     vFolderInfo = objFolder.ShortPath           'MS-DOS�`���p�X
        Case 9:     vFolderInfo = objFolder.DateCreated         '�쐬����
        Case 10:    vFolderInfo = objFolder.DateLastAccessed    '�A�N�Z�X����
        Case 11:    vFolderInfo = objFolder.DateLastModified    '�X�V����
        Case 12:    vFolderInfo = objFolder.Attributes          '����
        Case Else:  GetFolderInfo = False
    End Select
End Function
    Private Sub Test_GetFolderInfo()
        Dim sBuf As String
        Dim bRet As Boolean
        Dim vFolderInfo As Variant
        sBuf = ""
        Dim sTrgtPath As String
        sTrgtPath = "C:\codes\vbs\lib"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFolderInfo(sTrgtPath, 1, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C�����F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 2, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���T�C�Y�F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 3, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C����ށF" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 4, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���i�[��h���C�u�F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 5, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C���p�X�F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 6, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �e�t�H���_�F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 7, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  MS-DOS�`���t�@�C�����F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 8, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  MS-DOS�`���p�X�F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 9, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �쐬�����F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 10, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �A�N�Z�X�����F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 11, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �X�V�����F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 12, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �����F" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 13, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �F" & vFolderInfo
        sTrgtPath = "C:\codes\vbs\libs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFolderInfo(sTrgtPath, 1, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  �t�@�C�����F" & vFolderInfo
        MsgBox sBuf
    End Sub

'�t�H���_���́u�t�@�C�����v�u�����v���ݒ�\
'�������A�ȉ��̃��\�b�h�ɂĕύX�\�Ȃ��߁A���߂Ċ֐���`���Ȃ�
'  �t�@�C�����F objFSO.MoveFolder
'  �����F objFSO.GetFolder( "C:\codes" ).Attributes
'Public Function SetFolderInfo( _
'   ByVal sTrgtPath As String, _
'   ByVal lSetInfoType As Long, _
'   ByVal vFileInfo As Variant _
') As Boolean
'End Function

' ==================================================================
' = �T�v    �t�@�C���ڍ׏��擾
' = ����    sTrgtPath       String      [in]    �t�@�C���p�X
' = ����    lGetInfoType    Long        [in]    �擾����ʔԍ�(��)
' = ����    vFileInfoValue  Variant     [out]   �t�@�C���ڍ׏��
' = ����    vFileInfoTitle  Variant     [out]   �t�@�C���ڍ׏��^�C�g��
' = �ߒl                    Boolean             �擾����
' = �o��    (��)�擾�ł�����͂n�r�̃o�[�W�����ɂ���ĈقȂ�B
' =             ���O�� Exec_GetDetailsOfGetDetailsOf() �����s���āA
' =             �擾�ł�������m�F���Ă������ƁB
' =             �Ȃ��AlGetInfoType �� Folder �I�u�W�F�N�g GetDetailsOf()
' =             �v���p�e�B�̗v�f�ԍ��ɑΉ�����B
' =             ���蓖�Ă��Ă��Ȃ��擾����ʔԍ����w�肵���ꍇ�A
' =             �擾���� False ��ԋp����B
' ==================================================================
Public Function GetFileDetailInfo( _
    ByVal sTrgtPath As String, _
    ByVal lGetInfoType As Long, _
    ByRef vFileInfoValue As Variant, _
    ByRef vFileInfoTitle As Variant _
) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(sTrgtPath) Then
        'Do Nothing
    Else
        vFileInfoValue = ""
        GetFileDetailInfo = False
        Exit Function
    End If
    
    Dim sTrgtFolderPath As String
    Dim sTrgtFileName As String
    sTrgtFolderPath = Mid(sTrgtPath, 1, InStrRev(sTrgtPath, "\") - 1)
    sTrgtFileName = Mid(sTrgtPath, InStrRev(sTrgtPath, "\") + 1, Len(sTrgtPath))
    
    Dim objFolder As Object
    Set objFolder = CreateObject("Shell.Application").Namespace(sTrgtFolderPath & "\")
    Dim objFile As Object
    Set objFile = objFolder.ParseName(sTrgtFileName)
    
    vFileInfoValue = objFolder.GetDetailsOf(objFile, lGetInfoType)
    vFileInfoTitle = objFolder.GetDetailsOf("", lGetInfoType)
    If vFileInfoTitle = "" Then
        GetFileDetailInfo = False
    Else
        GetFileDetailInfo = True
    End If
End Function
    Private Sub Test_GetFileDetailInfo()
        Dim sBuf As String
        Dim bRet As Boolean
        Dim vFileInfoValue As Variant
        Dim vFileInfoTitle As Variant
        sBuf = ""
        Dim sTrgtPath As String
        sTrgtPath = "C:\codes\vbs\lib\FileSystem.vbs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFileDetailInfo(sTrgtPath, 1, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue
        bRet = GetFileDetailInfo(sTrgtPath, 2, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue
        bRet = GetFileDetailInfo(sTrgtPath, 3, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue
        bRet = GetFileDetailInfo(sTrgtPath, 4, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue
        bRet = GetFileDetailInfo(sTrgtPath, 52, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "�F" & vFileInfoValue
        MsgBox sBuf
    End Sub

'GetDetailsOf()�̏ڍ׏��i�v�f�ԍ��A�^�C�g�����A�^���A�f�[�^�j���擾����
Public Sub Exec_GetDetailsOfGetDetailsOf()
    Dim sTrgtFilePath As String
    Dim sLogFilePath As String
    sTrgtFilePath = "Z:\300_Musics\200_Reggae@Jamaica\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3"
    sLogFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\track_title_names.txt"
    
    Call GetDetailsOfGetDetailsOf(sTrgtFilePath, sLogFilePath)
    
    CreateObject("WScript.Shell").Run "%comspec% /c """ & sLogFilePath & """"
End Sub

'GetDetailsOf()�̏ڍ׏��i�v�f�ԍ��A�^�C�g�����A�^���A�f�[�^�j���擾����
Private Function GetDetailsOfGetDetailsOf( _
    ByVal sTrgtFilePath As String, _
    ByVal sLogFilePath As String _
)
    Dim sTrgtFolderPath As String
    Dim sTrgtFileName As String
    sTrgtFolderPath = Mid(sTrgtFilePath, 1, InStrRev(sTrgtFilePath, "\") - 1)
    sTrgtFileName = Mid(sTrgtFilePath, InStrRev(sTrgtFilePath, "\") + 1, Len(sTrgtFilePath))
    
    Dim objFolder As Object
    Set objFolder = CreateObject("Shell.Application").Namespace(sTrgtFolderPath & "\")
    Dim objFile As Object
    Set objFile = objFolder.ParseName(sTrgtFileName)
    
    Dim objTxtFile As Object
    Set objTxtFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(sLogFilePath, 2, True)
    objTxtFile.WriteLine "[Idx] " & Chr(9) & "[TypeName]" & Chr(9) & "[Title]"
    Dim i As Long
    For i = 0 To 400
        objTxtFile.WriteLine _
            i & Chr(9) & _
            TypeName(objFolder.GetDetailsOf(objFile, i)) & Chr(9) & _
            objFolder.GetDetailsOf("", i)
    Next i
    objTxtFile.Close
End Function
