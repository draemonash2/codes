Attribute VB_Name = "AlignMeasurementData"
Sub AlignMeasurementData()
    Dim ret As String
    Dim strInputCsvPath As String
    Dim strExportSheetName As String
    Dim strAlignSheetName As String
    
    strInputCsvPath = "D:\TANITA\GRAPHV1\DATA\DATA1.CSV"
    strExportSheetName = "DATA1"
    strAlignSheetName = "����\"
    
    ' CSV �ǂݍ���
    ret = openCsv(strInputCsvPath, strExportSheetName)
    
    ' �ǂݍ��݃f�[�^���` (����\�쐬)
    ret = alignData(strExportSheetName, strAlignSheetName)
    
    ' �g��(�i�q)�ݒ�
    ret = setGridRule(strAlignSheetName)
End Sub

' CSV ����
Function openCsv(strCsvPath As String, strExportSheetName As String) As String
    Dim strGetLine As String
    Dim arrGetLines As Variant
    Dim longSearchRow As Long
    Open strCsvPath For Input As #1
        longSearchRow = 1
        ' �ŏI�s�܂Ŏ擾
        Do Until EOF(1)
            Line Input #1, strGetLine
            arrGetLines = Split(strGetLine, ",")
            ' �ŏI��܂Ŏ擾����
            For longSearchColumn2 = 1 To UBound(arrGetLines)
                Worksheets(strExportSheetName).Cells(longSearchRow, longSearchColumn2).Value = arrGetLines(longSearchColumn2)
            Next longSearchColumn2
            longSearchRow = longSearchRow + 1
        Loop
    Close #1
    openCsv = "0"
End Function

' �f�[�^���`
Function alignData(strExportSheetName As String, strAlignSheetName As String) As String
    Dim strSearchData(1 To 12) As String
    
    ' �ǂݍ��݃f�[�^���` (����\�쐬)
    strSearchData(1) = "DT"    ' �����
    strSearchData(2) = "Ti"    ' ���莞��
    strSearchData(3) = "Hm"    ' �g�� (cm)
    strSearchData(4) = "Wk"    ' �̏d (kg)
    strSearchData(5) = "MI"    ' BMI
    strSearchData(6) = "FW"    ' �S�g�̎��b�� (%)
    strSearchData(7) = "mW"    ' �S�g�ؓ��� (kg)
    strSearchData(8) = "bW"    ' ���荜��(kg)
    strSearchData(9) = "IF"    ' �������b���x��
    strSearchData(10) = "rB"   ' ��b��ӗ� (kcal day)
    strSearchData(11) = "rA"   ' �̓��N�� (��)
    strSearchData(12) = "ww"   ' �̐����� (%)
    
    Worksheets(strExportSheetName).Activate
    
    ' �ŏI�s�܂ŌJ��Ԃ�
    For intSearchRow = 1 To (Range("A1").End(xlDown).Row)
        ' �����ʕ��J��Ԃ�
        For intSearchCnt = 1 To 12
            Worksheets(strAlignSheetName).Cells(intSearchRow + 1, intSearchCnt).Value = getMeasurementData(strSearchData(intSearchCnt), intSearchRow)
        Next intSearchCnt
    Next intSearchRow
    
    alignData = "0"
    
End Function


' ������擾
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

' �r��(�i�q)�ݒ�
Function setGridRule(strAlignSheetName As String) As String
    Dim MaxRow As String
    Dim MaxCol As String
    
    Worksheets(strAlignSheetName).Activate
    MaxRow = Range("A1").End(xlDown).Row
    MaxCol = Range("A1").End(xlToRight).Column
    
    ' �ŏI�s�܂ŌJ��Ԃ�
    For intRowCnt = 1 To MaxRow
        ' �����ʕ��J��Ԃ�
        For intColCnt = 1 To MaxCol
            Cells(intRowCnt, intColCnt).Borders.LineStyle = xlContinuous
        Next intColCnt
    Next intRowCnt
    
    setGridRule = "0"
End Function
