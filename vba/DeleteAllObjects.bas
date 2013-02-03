Attribute VB_Name = "DeleteAllObjects"
Sub DeleteAllObjects()
    Dim i As Long
    Dim j As Long
    
    For j = Worksheets.Count To 1 Step -1
        With Worksheets(j)
            For i = .ChartObjects.Count To 1 Step -1
                .ChartObjects(i).Delete
            Next i
        End With
    Next j
End Sub
