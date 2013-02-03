Attribute VB_Name = "AlignMeasurementData"
Sub AlignMeasurementData()
    Dim ret As String
    Dim strInputCsvPath As String
    Dim strExportSheetName As String
    Dim strAlignSheetName As String
    
    strInputCsvPath = "D:\TANITA\GRAPHV1\DATA\DATA1.CSV"
    strExportSheetName = "DATA1"
    strAlignSheetName = "測定表"
    
    ' CSV 読み込み
    ret = openCsv(strInputCsvPath, strExportSheetName)
    
    ' 読み込みデータ整形 (測定表作成)
    ret = alignData(strExportSheetName, strAlignSheetName)
    
    ' 枠線(格子)設定
    ret = setGridRule(strAlignSheetName)
End Sub

' CSV 入力
Function openCsv(strCsvPath As String, strExportSheetName As String) As String
    Dim strGetLine As String
    Dim arrGetLines As Variant
    Dim longSearchRow As Long
    Open strCsvPath For Input As #1
        longSearchRow = 1
        ' 最終行まで取得
        Do Until EOF(1)
            Line Input #1, strGetLine
            arrGetLines = Split(strGetLine, ",")
            ' 最終列まで取得する
            For longSearchColumn2 = 1 To UBound(arrGetLines)
                Worksheets(strExportSheetName).Cells(longSearchRow, longSearchColumn2).Value = arrGetLines(longSearchColumn2)
            Next longSearchColumn2
            longSearchRow = longSearchRow + 1
        Loop
    Close #1
    openCsv = "0"
End Function

' データ整形
Function alignData(strExportSheetName As String, strAlignSheetName As String) As String
    Dim strSearchData(1 To 12) As String
    
    ' 読み込みデータ整形 (測定表作成)
    strSearchData(1) = "DT"    ' 測定日
    strSearchData(2) = "Ti"    ' 測定時刻
    strSearchData(3) = "Hm"    ' 身長 (cm)
    strSearchData(4) = "Wk"    ' 体重 (kg)
    strSearchData(5) = "MI"    ' BMI
    strSearchData(6) = "FW"    ' 全身体脂肪率 (%)
    strSearchData(7) = "mW"    ' 全身筋肉量 (kg)
    strSearchData(8) = "bW"    ' 推定骨量(kg)
    strSearchData(9) = "IF"    ' 内臓脂肪レベル
    strSearchData(10) = "rB"   ' 基礎代謝量 (kcal day)
    strSearchData(11) = "rA"   ' 体内年齢 (才)
    strSearchData(12) = "ww"   ' 体水分量 (%)
    
    Worksheets(strExportSheetName).Activate
    
    ' 最終行まで繰り返す
    For intSearchRow = 1 To (Range("A1").End(xlDown).Row)
        ' 測定種別分繰り返す
        For intSearchCnt = 1 To 12
            Worksheets(strAlignSheetName).Cells(intSearchRow + 1, intSearchCnt).Value = getMeasurementData(strSearchData(intSearchCnt), intSearchRow)
        Next intSearchCnt
    Next intSearchRow
    
    alignData = "0"
    
End Function


' 文字列取得
Function getMeasurementData(strSearchData As String, ByVal intSearchRow As Integer) As String
    Dim ret2 As String
    
    ret2 = "error!"
    
    For strSearchColumn = 1 To 46
        Cells(intSearchRow, strSearchColumn).Select
        If ActiveCell.Value = strSearchData Then
            ret2 = (Cells(intSearchRow, strSearchColumn + 1))
            Exit For
        Else
            ' None
        End If
    Next strSearchColumn
    
    ret2 = Replace(ret2, """", "")
    
    getMeasurementData = ret2
End Function

' 罫線(格子)設定
Function setGridRule(strAlignSheetName As String) As String
    Dim MaxRow As String
    Dim MaxCol As String
    
    Worksheets(strAlignSheetName).Activate
    MaxRow = Range("A1").End(xlDown).Row
    MaxCol = Range("A1").End(xlToRight).Column
    
    ' 最終行まで繰り返す
    For intRowCnt = 1 To MaxRow
        ' 測定種別分繰り返す
        For intColCnt = 1 To MaxCol
            Cells(intRowCnt, intColCnt).Borders.LineStyle = xlContinuous
        Next intColCnt
    Next intRowCnt
    
    setGridRule = "0"
End Function
