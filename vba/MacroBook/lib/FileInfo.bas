Attribute VB_Name = "FileInfo"
Option Explicit

' file info libary v1.1

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

' ==================================================================
' = 概要    ファイル情報取得
' = 引数    sTrgtPath       String      [in]    ファイルパス
' = 引数    lGetInfoType    Long        [in]    取得情報種別 (※1)
' = 引数    vFileInfo       Variant     [out]   ファイル情報 (※1)
' = 戻値                    Boolean             取得結果
' = 覚書    以下、参照。
' =     (※1) ファイル情報
' =         [引数]  [説明]                  [プロパティ名]      [データ型]              [Get/Set]   [出力例]
' =         1       ファイル名              Name                vbString    文字列型    Get/Set     03 Ride Featuring Tony Matterhorn.MP3
' =         2       ファイルサイズ          Size                vbLong      長整数型    Get         4286923
' =         3       ファイル種類            Type                vbString    文字列型    Get         MPEG layer 3
' =         4       ファイル格納先ドライブ  Drive               vbString    文字列型    Get         Z:
' =         5       ファイルパス            Path                vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =         6       親フォルダ              ParentFolder        vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         7       MS-DOS形式ファイル名    ShortName           vbString    文字列型    Get         03 Ride Featuring Tony Matterhorn.MP3
' =         8       MS-DOS形式パス          ShortPath           vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =         9       作成日時                DateCreated         vbDate      日付型      Get         2015/08/19 0:54:45
' =         10      アクセス日時            DateLastAccessed    vbDate      日付型      Get         2016/10/14 6:00:30
' =         11      更新日時                DateLastModified    vbDate      日付型      Get         2016/10/14 6:00:30
' =         12      属性                    Attributes          vbLong      長整数型    (※2)       32
' =     (※2) 属性
' =         [値]                [説明]                                      [属性名]    [Get/Set]
' =         1  （0b00000001）   読み取り専用ファイル                        ReadOnly    Get/Set
' =         2  （0b00000010）   隠しファイル                                Hidden      Get/Set
' =         4  （0b00000100）   システム・ファイル                          System      Get/Set
' =         8  （0b00001000）   ディスクドライブ・ボリューム・ラベル        Volume      Get
' =         16 （0b00010000）   フォルダ／ディレクトリ                      Directory   Get
' =         32 （0b00100000）   前回のバックアップ以降に変更されていれば1   Archive     Get/Set
' =         64 （0b01000000）   リンク／ショートカット                      Alias       Get
' =         128（0b10000000）   圧縮ファイル                                Compressed  Get
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
        Case 1:     vFileInfo = objFile.Name                'ファイル名
        Case 2:     vFileInfo = objFile.Size                'ファイルサイズ
        Case 3:     vFileInfo = objFile.Type                'ファイル種類
        Case 4:     vFileInfo = objFile.Drive               'ファイル格納先ドライブ
        Case 5:     vFileInfo = objFile.Path                'ファイルパス
        Case 6:     vFileInfo = objFile.ParentFolder        '親フォルダ
        Case 7:     vFileInfo = objFile.ShortName           'MS-DOS形式ファイル名
        Case 8:     vFileInfo = objFile.ShortPath           'MS-DOS形式パス
        Case 9:     vFileInfo = objFile.DateCreated         '作成日時
        Case 10:    vFileInfo = objFile.DateLastAccessed    'アクセス日時
        Case 11:    vFileInfo = objFile.DateLastModified    '更新日時
        Case 12:    vFileInfo = objFile.Attributes          '属性
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
        bRet = GetFileInfo(sTrgtPath, 1, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイル名：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 2, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイルサイズ：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 3, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイル種類：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 4, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイル格納先ドライブ：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 5, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイルパス：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 6, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  親フォルダ：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 7, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  MS-DOS形式ファイル名：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 8, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  MS-DOS形式パス：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 9, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  作成日時：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 10, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  アクセス日時：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 11, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  更新日時：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 12, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  属性：" & vFileInfo
        bRet = GetFileInfo(sTrgtPath, 13, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  ：" & vFileInfo
        sTrgtPath = "C:\codes\vbs\lib\dummy.vbs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFileInfo(sTrgtPath, 1, vFileInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイル名：" & vFileInfo
        MsgBox sBuf
    End Sub

'ファイル情報は「ファイル名」「属性」が設定可能
'しかし、以下のメソッドにて変更可能なため、改めて関数定義しない
'  ファイル名： objFSO.MoveFile
'  属性： objFSO.GetFile( "C:\codes\a.txt" ).Attributes
'Public Function SetFileInfo( _
'   ByVal sTrgtPath As String, _
'   ByVal lSetInfoType As Long, _
'   ByVal vFileInfo As Variant _
') As Boolean
'End Function

' ==================================================================
' = 概要    フォルダ情報取得
' = 引数    sTrgtPath       String      [in]    フォルダパス
' = 引数    lGetInfoType    Long        [in]    取得情報種別 (※1)
' = 引数    vFolderInfo     Variant     [out]   フォルダ情報 (※1)
' = 戻値                    Boolean             取得結果
' = 覚書    以下、参照。
' =     (※1) フォルダ情報
' =         [引数]  [説明]                  [プロパティ名]      [データ型]              [Get/Set]   [出力例]
' =         1       フォルダ名              Name                vbString    文字列型    Get/Set     Sacrifice
' =         2       フォルダサイズ          Size                vbLong      長整数型    Get         80613775
' =         3       ファイル種類            Type                vbString    文字列型    Get         ファイル フォルダー
' =         4       ファイル格納先ドライブ  Drive               vbString    文字列型    Get         Z:
' =         5       フォルダパス            Path                vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         6       ルート フォルダ         IsRootFolder        vbBoolean   ブール型    Get         False
' =         7       MS-DOS形式ファイル名    ShortName           vbString    文字列型    Get         Sacrifice
' =         8       MS-DOS形式パス          ShortPath           vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =         9       作成日時                DateCreated         vbDate      日付型      Get         2015/08/19 0:54:44
' =         10      アクセス日時            DateLastAccessed    vbDate      日付型      Get         2015/08/19 0:54:44
' =         11      更新日時                DateLastModified    vbDate      日付型      Get         2015/04/18 3:38:36
' =         12      属性                    Attributes          vbLong      長整数型    (※2)       16
' =     (※2) 属性
' =         [値]                [説明]                                      [属性名]    [Get/Set]
' =         1  （0b00000001）   読み取り専用ファイル                        ReadOnly    Get/Set
' =         2  （0b00000010）   隠しファイル                                Hidden      Get/Set
' =         4  （0b00000100）   システム・ファイル                          System      Get/Set
' =         8  （0b00001000）   ディスクドライブ・ボリューム・ラベル        Volume      Get
' =         16 （0b00010000）   フォルダ／ディレクトリ                      Directory   Get
' =         32 （0b00100000）   前回のバックアップ以降に変更されていれば1   Archive     Get/Set
' =         64 （0b01000000）   リンク／ショートカット                      Alias       Get
' =         128（0b10000000）   圧縮ファイル                                Compressed  Get
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
        Case 1:     vFolderInfo = objFolder.Name                'フォルダ名
        Case 2:     vFolderInfo = objFolder.Size                'フォルダサイズ
        Case 3:     vFolderInfo = objFolder.Type                'ファイル種類
        Case 4:     vFolderInfo = objFolder.Drive               'ファイル格納先ドライブ
        Case 5:     vFolderInfo = objFolder.Path                'フォルダパス
        Case 6:     vFolderInfo = objFolder.IsRootFolder        'ルート フォルダ
        Case 7:     vFolderInfo = objFolder.ShortName           'MS-DOS形式ファイル名
        Case 8:     vFolderInfo = objFolder.ShortPath           'MS-DOS形式パス
        Case 9:     vFolderInfo = objFolder.DateCreated         '作成日時
        Case 10:    vFolderInfo = objFolder.DateLastAccessed    'アクセス日時
        Case 11:    vFolderInfo = objFolder.DateLastModified    '更新日時
        Case 12:    vFolderInfo = objFolder.Attributes          '属性
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
        bRet = GetFolderInfo(sTrgtPath, 1, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイル名：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 2, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイルサイズ：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 3, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイル種類：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 4, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイル格納先ドライブ：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 5, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイルパス：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 6, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  親フォルダ：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 7, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  MS-DOS形式ファイル名：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 8, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  MS-DOS形式パス：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 9, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  作成日時：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 10, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  アクセス日時：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 11, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  更新日時：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 12, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  属性：" & vFolderInfo
        bRet = GetFolderInfo(sTrgtPath, 13, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  ：" & vFolderInfo
        sTrgtPath = "C:\codes\vbs\libs"
        sBuf = sBuf & vbNewLine & sTrgtPath
        bRet = GetFolderInfo(sTrgtPath, 1, vFolderInfo): sBuf = sBuf & vbNewLine & bRet & "  ファイル名：" & vFolderInfo
        MsgBox sBuf
    End Sub

'フォルダ情報は「ファイル名」「属性」が設定可能
'しかし、以下のメソッドにて変更可能なため、改めて関数定義しない
'  ファイル名： objFSO.MoveFolder
'  属性： objFSO.GetFolder( "C:\codes" ).Attributes
'Public Function SetFolderInfo( _
'   ByVal sTrgtPath As String, _
'   ByVal lSetInfoType As Long, _
'   ByVal vFileInfo As Variant _
') As Boolean
'End Function

' ==================================================================
' = 概要    ファイル詳細情報取得
' = 引数    sTrgtPath       String      [in]    ファイルパス
' = 引数    lGetInfoType    Long        [in]    取得情報種別番号(※)
' = 引数    vFileInfoValue  Variant     [out]   ファイル詳細情報
' = 引数    vFileInfoTitle  Variant     [out]   ファイル詳細情報タイトル
' = 戻値                    Boolean             取得結果
' = 覚書    (※)取得できる情報はＯＳのバージョンによって異なる。
' =             事前に Exec_GetDetailsOfGetDetailsOf() を実行して、
' =             取得できる情報を確認しておくこと。
' =             なお、lGetInfoType は Folder オブジェクト GetDetailsOf()
' =             プロパティの要素番号に対応する。
' =             割り当てられていない取得情報種別番号を指定した場合、
' =             取得結果 False を返却する。
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
        bRet = GetFileDetailInfo(sTrgtPath, 1, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue
        bRet = GetFileDetailInfo(sTrgtPath, 2, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue
        bRet = GetFileDetailInfo(sTrgtPath, 3, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue
        bRet = GetFileDetailInfo(sTrgtPath, 4, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue
        bRet = GetFileDetailInfo(sTrgtPath, 52, vFileInfoValue, vFileInfoTitle): sBuf = sBuf & vbNewLine & bRet & "  " & vFileInfoTitle & "：" & vFileInfoValue
        MsgBox sBuf
    End Sub

'GetDetailsOf()の詳細情報（要素番号、タイトル情報、型名、データ）を取得する
Public Sub Exec_GetDetailsOfGetDetailsOf()
    Dim sTrgtFilePath As String
    Dim sLogFilePath As String
    sTrgtFilePath = "Z:\300_Musics\200_Reggae@Jamaica\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3"
    sLogFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\track_title_names.txt"
    
    Call GetDetailsOfGetDetailsOf(sTrgtFilePath, sLogFilePath)
    
    CreateObject("WScript.Shell").Run "%comspec% /c """ & sLogFilePath & """"
End Sub

'GetDetailsOf()の詳細情報（要素番号、タイトル情報、型名、データ）を取得する
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
