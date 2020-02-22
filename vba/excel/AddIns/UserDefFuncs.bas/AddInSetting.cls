VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddInSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'AddInSetting.cls v1.0.0a

' ==================================================================
' = 概要    アドイン用の設定シートを管理するクラス
' ==================================================================

' ==================================================================
' = 設定
' ==================================================================
Private Const SHEET_NAME = "setting"
Private Const CLM_KEY As Long = 1
Private Const CLM_VALUE As Long = 2

' ==================================================================
' = 概要    設定数を返却する
' = 引数    なし
' = 戻値                    Long    キー数
' = 覚書    なし
' = 依存    なし
' = 所属    AddInSetting.cls
' ==================================================================
Public Property Get Count() As Long
    With ThisWorkbook.Sheets(SHEET_NAME)
        Dim lLastRow As Long
        lLastRow = .Cells(.Rows.Count, CLM_KEY).End(xlUp).Row
        If .Cells(lLastRow, CLM_KEY).Value = "" Then
            Count = 0
        Else
            Count = lLastRow
        End If
    End With
End Property

' ==================================================================
' = 概要    設定を追加する
' = 引数    sKey            String      [in]    設定キー
' = 引数    sValue          String      [in]    設定値
' = 戻値                    Boolean             追加結果(True固定)
' = 覚書    なし
' = 依存    なし
' = 所属    AddInSetting.cls
' ==================================================================
Public Function Add( _
    ByVal sKey As String, _
    ByVal sValue As String _
) As Boolean
    With ThisWorkbook.Sheets(SHEET_NAME)
        Dim lLastRow As Long
        lLastRow = .Cells(.Rows.Count, CLM_KEY).End(xlUp).Row
        
        '検索
        Dim bIsMatch As Boolean
        bIsMatch = False
        Dim lRowIdx As Long
        For lRowIdx = 1 To lLastRow
            If .Cells(lRowIdx, CLM_KEY).Value = sKey Then
                bIsMatch = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lRowIdx
        
        '追加先決定
        Dim lAddRow As Long
        If bIsMatch = True Then
            lAddRow = lRowIdx
        Else
            If lLastRow = 1 Then
                If .Cells(lLastRow, CLM_KEY).Value = "" Then
                    lAddRow = 1
                Else
                    lAddRow = 2
                End If
            Else
                lAddRow = lLastRow + 1
            End If
        End If
        
        '更新/追加
        .Cells(lAddRow, CLM_KEY).Value = sKey
        .Cells(lAddRow, CLM_VALUE).Value = sValue
        
        Add = True 'True 以外は返さない
    End With
End Function

' ==================================================================
' = 概要    設定を更新する
' = 引数    sKey            String      [in]    設定キー
' = 引数    sValue          String      [in]    設定値
' = 戻値                    Boolean             更新結果(True:更新成功,False:キー未存在)
' = 覚書    設定キーが存在しない場合は何もしない
' = 依存    なし
' = 所属    AddInSetting.cls
' ==================================================================
Public Function Update( _
    ByVal sKey As String, _
    ByVal sValue As String _
) As Boolean
    With ThisWorkbook.Sheets(SHEET_NAME)
        Dim lLastRow As Long
        lLastRow = .Cells(.Rows.Count, CLM_KEY).End(xlUp).Row
        
        '検索
        Dim bIsMatch As Boolean
        bIsMatch = False
        Dim lRowIdx As Long
        For lRowIdx = 1 To lLastRow
            If .Cells(lRowIdx, CLM_KEY).Value = sKey Then
                bIsMatch = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lRowIdx
        
        '更新
        If bIsMatch = True Then
            .Cells(lRowIdx, CLM_KEY).Value = sKey
            .Cells(lRowIdx, CLM_VALUE).Value = sValue
            Update = True
        Else
            Update = False
        End If
    End With
End Function

' ==================================================================
' = 概要    設定を削除する
' = 引数    sKey            String      [in]    設定キー
' = 戻値                    Boolean             削除結果(True:削除成功,False:キー未存在)
' = 覚書    設定キーが存在しない場合は何もしない
' = 依存    なし
' = 所属    AddInSetting.cls
' ==================================================================
Public Function Delete( _
    ByVal sKey As String _
) As Boolean
    With ThisWorkbook.Sheets(SHEET_NAME)
        Delete = False
        Dim lLastRow As Long
        lLastRow = .Cells(.Rows.Count, CLM_KEY).End(xlUp).Row
        Dim lRowIdx As Long
        For lRowIdx = 1 To lLastRow
            If .Cells(lRowIdx, CLM_KEY).Value = sKey Then
                .Cells(lRowIdx, CLM_KEY).EntireRow.Delete Shift:=xlShiftUp
                Delete = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lRowIdx
    End With
End Function

' ==================================================================
' = 概要    設定キーをもとに設定値を検索する
' = 引数    sKey            String      [in]    設定キー
' = 引数    sValue          String      [out]   設定値
' = 戻値                    Boolean             検索結果(True:検索成功,False:キー未存在)
' = 覚書    なし
' = 依存    なし
' = 所属    AddInSetting.cls
' ==================================================================
Public Function SearchWithKey( _
    ByVal sKey As String, _
    ByRef sValue As String _
) As Boolean
    With ThisWorkbook.Sheets(SHEET_NAME)
        sValue = ""
        SearchWithKey = False
        Dim lLastRow As Long
        lLastRow = .Cells(.Rows.Count, CLM_KEY).End(xlUp).Row
        Dim lRowIdx As Long
        For lRowIdx = 1 To lLastRow
            If .Cells(lRowIdx, CLM_KEY).Value = sKey Then
                sValue = .Cells(lRowIdx, CLM_VALUE).Value
                SearchWithKey = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lRowIdx
    End With
End Function

' ==================================================================
' = 概要    設定インデックスをもとに設定値を検索する
' = 引数    lIdx            Long        [in]    設定インデックス(設定シートの行番号)
' = 引数    sKey            String      [out]   設定キー
' = 引数    sValue          String      [out]   設定値
' = 戻値                    Boolean             検索結果(True:検索成功,False:キー未存在)
' = 覚書    なし
' = 依存    なし
' = 所属    AddInSetting.cls
' ==================================================================
Public Function SearchWithIdx( _
    ByVal lIdx As Long, _
    ByRef sKey As String, _
    ByRef sValue As String _
) As Boolean
    With ThisWorkbook.Sheets(SHEET_NAME)
        Dim lLastRow As Long
        lLastRow = .Cells(.Rows.Count, CLM_KEY).End(xlUp).Row
        
        If lIdx >= 1 And lIdx <= lLastRow Then
            sKey = .Cells(lIdx, CLM_KEY).Value
            sValue = .Cells(lIdx, CLM_VALUE).Value
            SearchWithIdx = True
        Else
            sKey = ""
            sValue = ""
            SearchWithIdx = False
        End If
    End With
End Function

