Attribute VB_Name = "CopyPasteForVisibilityCell"
Sub CopyPasteForVisibilityCell()
    Worksheets("Sheet1").Activate
End Sub

'Sheet1�̃Z��B1�`D3�͈̔͂̍폜(�����ɂ�������ɃV�t�g)
Sub DeleteCell()
     Worksheets("Sheet1").Activate
     Worksheets("Sheet1").Range("A2:A5").EntireRow.Delete Shift:=xlShiftUp
End Sub

