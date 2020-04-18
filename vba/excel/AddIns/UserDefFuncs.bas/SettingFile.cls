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

'setting file class v1.0.0

' ******************************************************************
' * 概要    アドイン用の設定ファイルを管理するクラス
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
    Debug.Print gsSettingFilePath
    
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
            vKeyValue = Split(sLine, gsDelimiter)
            Debug.Assert UBound(vKeyValue) = 1 '区切り文字が不一致
            gdSettingItems.Add vKeyValue(0), vKeyValue(1)
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
    If gsSettingFilePath = "" Then
        FileSave = False
    Else
        Debug.Print gsSettingFilePath
        
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
    If gsSettingFilePath = "" Then
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
    If gsSettingFilePath = "" Then
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
    If gsSettingFilePath = "" Then
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




