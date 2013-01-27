Attribute VB_Name = "CopyPasteForVisibilityCell"
Sub CopyPasteForVisibilityCell()
    Worksheets("Sheet1").Activate
End Sub

'Sheet1のセルB1～D3の範囲の削除(引数により上方向にシフト)
Sub DeleteCell()
     Worksheets("Sheet1").Activate
     Worksheets("Sheet1").Range("A2:A5").EntireRow.Delete Shift:=xlShiftUp
End Sub

