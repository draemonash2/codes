Attribute VB_Name = "Func"
Option Explicit

' ==================================================================
' = 概要    指定した２つの範囲を比較して、完全一致かどうかを判定する
' = 引数    rTrgtRange01        Range   [in]  比較対象範囲１
' = 引数    rTrgtRange02        Range   [in]  比較対象範囲２
' = 引数    bCellPosCheckValid  Boolean [in]  セル位置チェック有効/無効
' = 戻値                        Boolean       比較結果
' = 覚書    以下のいずれかを満たす場合、False を返却する
' =           ・範囲内のセル数が不一致
' =           ・範囲内の行数が不一致
' =           ・範囲内の列数が不一致
' =           ・範囲内の各セルの値が不一致
' =           ・範囲内の開始セルと末尾セルのセル位置が不一致
' ==================================================================
Public Function DiffRange( _
    ByRef rTrgtRange01 As Range, _
    ByRef rTrgtRange02 As Range, _
    Optional bCellPosCheckValid As Boolean = False _
) As Boolean
    DiffRange = True
    If rTrgtRange01.Count = rTrgtRange02.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    If rTrgtRange01.Rows.Count = rTrgtRange02.Rows.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    If rTrgtRange01.Columns.Count = rTrgtRange02.Columns.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    Dim lIdx As Long
    For lIdx = 1 To rTrgtRange01.Count
        If rTrgtRange01(lIdx).Value = rTrgtRange02(lIdx).Value Then
            'Do Nothing
        Else
            DiffRange = False
            Exit Function
        End If
    Next lIdx
    If bCellPosCheckValid = True Then
        If rTrgtRange01(1).Row = rTrgtRange02(1).Row And _
           rTrgtRange01(1).Column = rTrgtRange02(1).Column And _
           rTrgtRange01(rTrgtRange01.Count).Row = rTrgtRange02(rTrgtRange02.Count).Row And _
           rTrgtRange01(rTrgtRange01.Count).Column = rTrgtRange02(rTrgtRange02.Count).Column Then
            'Do Nothing
        Else
            DiffRange = False
            Exit Function
        End If
    Else
        'Do Nothing
    End If
    
End Function
    Private Sub Test_DiffRange()
        Dim shDiff01 As Worksheet
        Dim shDiff02 As Worksheet
        Set shDiff01 = ThisWorkbook.Sheets("タグ一覧")
        Set shDiff02 = ThisWorkbook.Sheets("タグ一覧_ミラー")
        Debug.Print DiffRange( _
            shDiff01.Range( _
                shDiff01.Cells(4, 6), _
                shDiff01.Cells(4, 39) _
            ), _
            shDiff02.Range( _
                shDiff02.Cells(4, 6), _
                shDiff02.Cells(4, 39) _
            ) _
        )
    End Sub
