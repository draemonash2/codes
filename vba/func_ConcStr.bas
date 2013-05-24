Attribute VB_Name = "ConcStr"
Function ConcStr(rConcRange As Range, Optional sDelimitState As String) As Variant
    ' �ϐ���`
    Dim rCongRangeCnt As Range
    Dim sConcTxtBuf As String
    
    If rConcRange.Rows.Count = 1 Or rConcRange.Columns.Count = 1 Then
        For Each rCongRangeCnt In rConcRange
            sConcTxtBuf = sConcTxtBuf & sDelimitState & rCongRangeCnt.Value
        Next rCongRangeCnt
        
        ' ��؂蕶������
        If sDelimitState <> "" Then
            ConcStr = Mid$(sConcTxtBuf, 2)
        Else
            ConcStr = sConcTxtBuf
        End If
    Else
        ConcStr = CVErr(xlErrRef)  '�G���[�l
    End If
End Function

