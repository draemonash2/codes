Attribute VB_Name = "FileInfo"
Option Explicit

' file info libary v1.0

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

