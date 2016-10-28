Attribute VB_Name = "FileInfo"
Option Explicit

' file info libary v1.0

' CreateFile 関数
Private Declare Function CreateFile Lib "KERNEL32.DLL" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

' CloseHandle 関数
Private Declare Function CloseHandle Lib "KERNEL32.DLL" ( _
    ByVal hObject As Long _
) As Long

' LocalFileTimeToFileTime 関数
Private Declare Function LocalFileTimeToFileTime Lib "KERNEL32.DLL" ( _
    ByRef lpLocalFileTime As FileTime, _
    ByRef lpFileTime As FileTime _
) As Long

' SystemTimeToFileTime 関数
Private Declare Function SystemTimeToFileTime Lib "KERNEL32.DLL" ( _
    ByRef lpSystemTime As SystemTime, _
    ByRef lpFileTime As FileTime _
) As Long

' SetFileTime 関数
Private Declare Function SetFileTime Lib "KERNEL32.DLL" ( _
    ByVal cFile As Long, _
    ByRef lpCreationTime As FileTime, _
    ByRef lpLastAccessTime As FileTime, _
    ByRef lpLastWriteTime As FileTime _
) As Long

' SystemTime 構造体
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

' FileTime 構造体
Private Type FileTime
    LowDateTime As Long
    HighDateTime As Long
End Type

' 定数の定義
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const OPEN_EXISTING As Long = 3

' FileTime を取得する
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

' ファイルのハンドルを取得する
Private Function GetFileHandle(ByVal stFilePath As String) As Long
    GetFileHandle = CreateFile( _
        stFilePath, GENERIC_READ Or GENERIC_WRITE, _
        FILE_SHARE_READ, 0, OPEN_EXISTING, _
        FILE_ATTRIBUTE_NORMAL, 0 _
    )
End Function

' -------------------------------------------------------------------------------
' 作成日時を指定した日付と時間に設定します。
'
' @Param stFilePath 作成日時を設定するファイルまでのパス。
' @Param stCreateTime 作成日時に設定する日付と時間。
' -------------------------------------------------------------------------------
Public Function SetCreationTime( _
    ByVal stFilePath As String, _
    ByVal stCreateTime As String _
)
    ' FileTime を取得する
    Dim tFileTime As FileTime
    tFileTime = GetFileTime(CDate(stCreateTime))
    
    ' ファイルのハンドルを取得する
    Dim cFileHandle As Long
    cFileHandle = GetFileHandle(stFilePath)
    
    ' ファイルのハンドルが取得できた場合のみ「作成日時」を更新する
    If cFileHandle >= 0 Then
        Dim tNullable As FileTime
        
        Call SetFileTime(cFileHandle, tFileTime, tNullable, tNullable)
        Call CloseHandle(cFileHandle)
    End If
End Function

' -------------------------------------------------------------------------------
' 更新日時を指定した日付と時間に設定します。
'
' @Param stFilePath 更新日時を設定するファイルまでのパス。
' @Param stUpdateTime 更新日時に設定する日付と時間。
' -------------------------------------------------------------------------------
Public Function SetLastWriteTime( _
    ByVal stFilePath As String, _
    ByVal stUpdateTime As String _
)
    ' FileTime を取得する
    Dim tFileTime As FileTime
    tFileTime = GetFileTime(CDate(stUpdateTime))
    
    ' ファイルのハンドルを取得する
    Dim cFileHandle As Long
    cFileHandle = GetFileHandle(stFilePath)
    
    ' ファイルのハンドルが取得できた場合のみ「更新日時」を更新する
    If cFileHandle >= 0 Then
        Dim tNullable As FileTime
        
        Call SetFileTime(cFileHandle, tNullable, tNullable, tFileTime)
        Call CloseHandle(cFileHandle)
    End If
End Function

' -------------------------------------------------------------------------------
' アクセス日時を指定した日付と時間に設定します。
'
' @Param stFilePath アクセス日時を設定するファイルまでのパス。
' @Param stAccessTime アクセス日時に設定する日付と時間。
' -------------------------------------------------------------------------------
Public Function SetLastAccessTime( _
    ByVal stFilePath As String, _
    ByVal stAccessTime As String _
)
    ' FileTime を取得する
    Dim tFileTime As FileTime
    tFileTime = GetFileTime(CDate(stAccessTime))
    
    ' ファイルのハンドルを取得する
    Dim cFileHandle As Long
    cFileHandle = GetFileHandle(stFilePath)
    
    ' ファイルのハンドルが取得できた場合のみ「アクセス日時」を更新する
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

