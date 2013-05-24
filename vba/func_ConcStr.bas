Attribute VB_Name = "ConcStr"
Function ConcStr(rConcRange As Range, Optional sDelimitState As String) As Variant
    ' 変数定義
    Dim rCongRangeCnt As Range
    Dim sConcTxtBuf As String
    
    If rConcRange.Rows.Count = 1 Or rConcRange.Columns.Count = 1 Then
        For Each rCongRangeCnt In rConcRange
            sConcTxtBuf = sConcTxtBuf & sDelimitState & rCongRangeCnt.Value
        Next rCongRangeCnt
        
        ' 区切り文字判定
        If sDelimitState <> "" Then
            ConcStr = Mid$(sConcTxtBuf, 2)
        Else
            ConcStr = sConcTxtBuf
        End If
    Else
        ConcStr = CVErr(xlErrRef)  'エラー値
    End If
End Function

