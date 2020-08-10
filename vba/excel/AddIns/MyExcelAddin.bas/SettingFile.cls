VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'setting file class v1.2.2

' ******************************************************************
' * 概要    アドイン用の設定ファイルを管理するクラス
' *
' * 覚書    設定は文字列型として保存する。そのため以下に注意すること
' *           - Boolean値は文字列を指定すること(ex.Trueではなく"True")
' *           - 制御文字は制御文字を示す文字列を指定すること(ex.vbTabではなく"vbTab")
' ******************************************************************

' ******************************************************************
' * 設定
' ******************************************************************
Const sDELIMITER_INIT As String = vbTab

' ******************************************************************
' * 本処理
' ******************************************************************
Private gdSettingItems As Object
Private gsSettingFilePath As String
Private gsDelimiter As String

' ==================================================================
' = 概要    コンストラクタ
' ==================================================================
Private Sub Class_Initialize()
    Call Init
End Sub

' ==================================================================
' = 概要    デストラクタ
' ==================================================================
Private Sub Class_Terminate()
    Call DeInit
End Sub

' ==================================================================
' = 概要    初期化する
' ==================================================================
Private Sub Init()
    Set gdSettingItems = CreateObject("Scripting.Dictionary")
    gsSettingFilePath = ""
    gsDelimiter = sDELIMITER_INIT
End Sub

' ==================================================================
' = 概要    終了する
' ==================================================================
Private Sub DeInit()
    Set gdSettingItems = Nothing
    gsSettingFilePath = ""
    gsDelimiter = sDELIMITER_INIT
End Sub

' ==================================================================
' = 概要    ファイルを読み出す
' = 引数    sFilePath       String      [in]    ファイルパス
' = 引数    sDelimiter      String      [in]    デリミタ(省略可)
' = 戻値                    Boolean             読み出し結果
' = 覚書    ・「読み出すファイルの区切り文字」とsDelimiterを一致させること
' =           一致しない場合、処理を中断する。
' =         ・以下の場合、Falseを返却する
' =           - sFilePathが存在しない
' = 依存    なし
' ==================================================================
Public Function FileLoad( _
    ByVal sFilePath As String, _
    Optional ByVal sDelimiter As String = sDELIMITER_INIT _
) As Boolean
    gsDelimiter = sDelimiter
    
    gsSettingFilePath = sFilePath
    'Debug.Print gsSettingFilePath
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(gsSettingFilePath) Then
        Dim vFileLineAll() As String
        
        Open gsSettingFilePath For Input As #1
        FileLoad = False
        Do Until EOF(1)
            Dim vKeyValue As Variant
            Dim sLine As String
            Line Input #1, sLine
            If InStr(sLine, gsDelimiter) Then
                vKeyValue = Split(sLine, gsDelimiter)
                If UBound(vKeyValue) = 0 Then
                    gdSettingItems.Add vKeyValue(0), ""           '単一区切り文字(値なし)
                ElseIf UBound(vKeyValue) = 1 Then
                    gdSettingItems.Add vKeyValue(0), vKeyValue(1) '単一区切り文字(値あり)
                Else
                    Stop                                          '複数区切り文字
                End If
            Else
                Stop                                              '区切り文字なし
            End If
            FileLoad = True
        Loop
        Close #1
    Else
        FileLoad = False
    End If
End Function

' ==================================================================
' = 概要    ファイルを書き出す
' = 引数    なし
' = 戻値                    Boolean             書き出し結果
' = 覚書    ・以下の場合、Falseを返却する
' =           - 事前にFileLoad()が呼ばれていない
' = 依存    なし
' ==================================================================
Public Function FileSave() As Boolean
    If gdSettingItems Is Nothing Then
        FileSave = False
    Else
        'Debug.Print gsSettingFilePath
        
        Open gsSettingFilePath For Output As #1
        Dim vKey As Variant
        For Each vKey In gdSettingItems
            Print #1, vKey & gsDelimiter & gdSettingItems.Item(vKey)
        Next
        Close #1
        FileSave = True
    End If
End Function

' ==================================================================
' = 概要    設定を追加する
' = 引数    sKey            String      [in]    設定キー
' = 引数    sValue          String      [in]    設定値
' = 引数    bDoSave         String      [in]    ファイル保存実施有無(省略可)
' = 戻値                    Boolean             追加結果
' = 覚書    ・以下の場合、Falseを返却する
' =           - 事前にFileLoad()が呼ばれていない
' =           - 保存に失敗
' = 依存    Me/FileSave()
' ==================================================================
Public Function Add( _
    ByVal sKey As String, _
    ByVal sValue As String, _
    Optional ByVal bDoSave As Boolean = True _
) As Boolean
    If gdSettingItems Is Nothing Then
        Add = False
    Else
        '追加
        If gdSettingItems.Exists(sKey) Then
            gdSettingItems.Item(sKey) = sValue
        Else
            gdSettingItems.Add sKey, sValue
        End If
        
        'ファイル保存
        If bDoSave = True Then
            Add = FileSave()
        Else
            Add = True
        End If
    End If
End Function

' ==================================================================
' = 概要    設定を削除する
' = 引数    sKey            String      [in]    設定キー
' = 引数    bDoSave         String      [in]    ファイル保存実施有無(省略可)
' = 戻値                    Boolean             削除結果
' = 覚書    ・以下の場合、Falseを返却する
' =           - sKeyが存在しない
' =           - 事前にFileLoad()が呼ばれていない
' =           - 保存に失敗
' = 依存    Me/FileSave()
' ==================================================================
Public Function Delete( _
    ByVal sKey As String, _
    Optional ByVal bDoSave As Boolean = True _
) As Boolean
    If gdSettingItems Is Nothing Then
        Delete = False
    Else
        If gdSettingItems.Exists(sKey) Then
            '削除
            gdSettingItems.Remove (sKey)
            
            'ファイル保存
            If bDoSave = True Then
                Delete = FileSave()
            Else
                Delete = True
            End If
        Else
            Delete = False
        End If
    End If
End Function

' ==================================================================
' = 概要    設定を全削除する
' = 引数    bDoSave         String      [in]    ファイル保存実施有無(省略可)
' = 戻値                    Boolean             削除結果
' = 覚書    ・以下の場合、Falseを返却する
' =           - 事前にFileLoad()が呼ばれていない
' =           - 保存に失敗
' = 依存    Me/FileSave()
' ==================================================================
Public Function DeleteAll( _
    Optional ByVal bDoSave As Boolean = True _
) As Boolean
    If gdSettingItems Is Nothing Then
        DeleteAll = False
    Else
        '全削除
        gdSettingItems.RemoveAll
        
        'ファイル保存
        If bDoSave = True Then
            DeleteAll = FileSave()
        Else
            DeleteAll = True
        End If
    End If
End Function

' ==================================================================
' = 概要    設定の存在確認を行う
' = 引数    sKey            String      [in]    設定キー
' = 戻値                    Boolean             存在確認結果
' = 覚書    ・以下の場合、Falseを返却する
' =           - 事前にFileLoad()が呼ばれていない
' =           - sKeyが存在しない
' = 依存    なし
' ==================================================================
Public Function Exists( _
    ByVal sKey As String _
) As Boolean
    If gdSettingItems Is Nothing Then
        Exists = False
    Else
        If gdSettingItems.Exists(sKey) Then
            Exists = True
        Else
            Exists = False
        End If
    End If
End Function

' ==================================================================
' = 概要    設定数を取得する
' = 引数    なし
' = 戻値                    Long    設定数
' = 覚書    ・以下の場合、0を返却する
' =           - 事前にFileLoad()が呼ばれていない
' = 依存    なし
' ==================================================================
Public Property Get Count() As Long
    If gdSettingItems Is Nothing Then
        Count = 0
    Else
        Count = gdSettingItems.Count
    End If
End Property

' ==================================================================
' = 概要    設定値を取得する
' = 引数    sKey            String      [in]    設定キー
' = 引数    sValue          String      [out]   設定値
' = 戻値                    Boolean             取得結果
' = 覚書    ・以下の場合、Falseを返却する
' =           - 事前にFileLoad()が呼ばれていない
' =           - sKeyが存在しない
' = 依存    なし
' ==================================================================
Public Function Item( _
    ByVal sKey As String, _
    ByRef sValue As String _
) As Boolean
    If gdSettingItems.Exists(sKey) Then
        sValue = gdSettingItems.Item(sKey)
        Item = True
    Else
        sValue = ""
        Item = False
    End If
End Function

' ==================================================================
' = 概要    設定キーと設定値を全て取得する
' = 引数    なし
' = 戻値                    Object(Dictionary)  設定キー/値辞書
' = 覚書    なし
' = 依存    なし
' ==================================================================
Public Property Get AllItems() As Object
    Set AllItems = gdSettingItems
End Property

' ==================================================================
' = 概要    ファイルオープンから設定取得（＆なければ設定追加）を一括で行う
' = 引数    sFilePath       String      [in]    ファイルパス
' = 引数    sKey            String      [in]    設定キー
' = 引数    sValue          String      [out]   設定値
' = 引数    sInitValue      String      [in]    設定初期値
' = 引数    bDoSave         String      [in]    ファイル保存実施有無(省略可)
' = 戻値                    Boolean             取得結果
' = 覚書    ・ファイルオープン後、設定値を取得する。
' =           設定値が存在しない場合、初期値として設定値を更新する。
' =         ・以下の場合、Falseを返却する
' =           - sFilePathが存在しない
' =           - sKeyが存在しない
' = 依存    なし
' ==================================================================
Public Function ReadItemFromFile( _
    ByVal sFilePath As String, _
    ByVal sKey As String, _
    ByRef sValue As String, _
    ByVal sInitValue As String, _
    Optional ByVal bDoSave As Boolean = True _
) As Boolean
    Call Init
    
    '格納先フォルダ作成
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Call CreateDirectry(objFSO.GetParentFolderName(sFilePath))
    
    '設定ファイル読み出し
    Dim bExistFile As Boolean
    Dim bExistItem As Boolean
    bExistFile = Me.FileLoad(sFilePath)
    
    '設定項目取得＆更新
    Dim sItem As String
    If bExistFile = True Then
        bExistItem = Me.Item(sKey, sItem)
        If bExistItem = True Then
            sValue = sItem
            'Call Me.Add(sKey, sValue, bDoSave)
            ReadItemFromFile = True
        Else
            sValue = sInitValue
            Call Me.Add(sKey, sValue, bDoSave)
            ReadItemFromFile = False
        End If
    Else
        sValue = sInitValue
        Call Me.Add(sKey, sValue, bDoSave)
        ReadItemFromFile = False
    End If
    
    Call DeInit
End Function

' ==================================================================
' = 概要    ファイルオープンから設定更新/追加を一括で行う
' = 引数    sFilePath       String      [in]    ファイルパス
' = 引数    sKey            String      [in]    設定キー
' = 引数    sValue          String      [in]    設定値
' = 戻値                    Boolean             取得結果
' = 覚書    ・ファイルオープン後、設定値を更新/追加する。
' =         ・以下の場合、Falseを返却する
' =           - sFilePathが存在しない
' =           - sKeyが存在しない
' = 依存    なし
' ==================================================================
Public Function WriteItemToFile( _
    ByVal sFilePath As String, _
    ByVal sKey As String, _
    ByVal sValue As String _
) As Boolean
    Call Init
    
    '格納先フォルダ作成
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Call CreateDirectry(objFSO.GetParentFolderName(sFilePath))
    
    '設定ファイル読み出し
    Dim bExistFile As Boolean
    Dim bExistItem As Boolean
    bExistFile = Me.FileLoad(sFilePath)
    bExistItem = Me.Add(sKey, sValue)
    If bExistFile = True And bExistItem = True Then
        WriteItemToFile = True
    Else
        WriteItemToFile = False
    End If
    
    Call DeInit
End Function

' ==================================================================
' = 概要    設定値変換用 文字列to真偽値変換
' = 引数    sValue          String      [in]    値(文字列)
' = 戻値                    Boolean             値(真偽値)
' = 覚書    なし
' = 依存    なし
' ==================================================================
Public Function ConvTypeStr2Bool( _
    ByVal sValue As String _
) As Boolean
    If sValue = "True" Then
        ConvTypeStr2Bool = True
    Else
        ConvTypeStr2Bool = False
    End If
End Function

' ==================================================================
' = 概要    設定値変換用 真偽値to文字列変換
' = 引数    bValue          Boolean     [in]    値(真偽値)
' = 戻値                    String              値(文字列)
' = 覚書    なし
' = 依存    なし
' ==================================================================
Public Function ConvTypeBool2Str( _
    ByVal bValue As Boolean _
) As String
    If bValue = True Then
        ConvTypeBool2Str = "True"
    Else
        ConvTypeBool2Str = "False"
    End If
End Function

' ==================================================================
' = 概要    設定値変換用 生値to制御文字 変換
' = 引数    sValue          String      [in]    値(文字列)
' = 戻値                    String              値(制御文字)
' = 覚書    なし
' = 依存    なし
' ==================================================================
Public Function ConvStrRaw2CntrlChr( _
    ByVal sValue As String _
) As String
    Select Case sValue
        Case "vbTab":     ConvStrRaw2CntrlChr = vbTab
        Case "vbCr":      ConvStrRaw2CntrlChr = vbCr
        Case "vbLf":      ConvStrRaw2CntrlChr = vbLf
        Case "vbNewLine": ConvStrRaw2CntrlChr = vbNewLine
        Case Else:        ConvStrRaw2CntrlChr = sValue
    End Select
End Function

' ==================================================================
' = 概要    設定値変換用 制御文字to生値 変換
' = 引数    sValue          String      [in]    値(制御文字)
' = 戻値                    String              値(文字列)
' = 覚書    なし
' = 依存    なし
' ==================================================================
Public Function ConvStrCntrlChr2Raw( _
    ByVal sValue As String _
) As String
    Select Case sValue
        Case vbTab:     ConvStrCntrlChr2Raw = "vbTab"
        Case vbCr:      ConvStrCntrlChr2Raw = "vbCr"
        Case vbLf:      ConvStrCntrlChr2Raw = "vbLf"
        Case vbNewLine: ConvStrCntrlChr2Raw = "vbNewLine"
        Case Else:      ConvStrCntrlChr2Raw = sValue
    End Select
End Function

' ==================================================================
' = 概要    ディレクトリを作成する。親ディレクトリも自動生成する。
' = 引数    sDirPath    String  [in]  フォルダパス
' = 戻値    なし
' = 覚書    フォルダが既に存在している場合は何もしない
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function CreateDirectry( _
    ByVal sDirPath As String _
)
    Dim sParentDir As String
    Dim oFileSys As Object
 
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
 
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
 
    '親ディレクトリが存在しない場合、再帰呼び出し
    If oFileSys.FolderExists(sParentDir) = False Then
        Call CreateDirectry(sParentDir)
    End If
 
    'ディレクトリ作成
    If oFileSys.FolderExists(sDirPath) = False Then
        oFileSys.CreateFolder sDirPath
    End If
 
    Set oFileSys = Nothing
End Function

