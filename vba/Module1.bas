Attribute VB_Name = "Module1"
Sub MeasurementData()
    Dim ret As String
    Dim strInputCsvPath As String
    Dim strExportSheetName As String
    Dim strAlignSheetName As String
    
    strInputCsvPath = "C:\Users\TatsuyaEndo\Dropbox\800_Note\DATA1.CSV"
    strExportSheetName = "DATA1"
    strAlignSheetName = "����\"
    
    ' CSV �ǂݍ���
    ret = openCsv(strInputCsvPath, strExportSheetName)
    
    ' �ǂݍ��݃f�[�^���` (����\�쐬)
    ret = alignData(strExportSheetName, strAlignSheetName)
   
End Sub

' CSV ���͂̃v���V�W��
Function openCsv(strCsvPath As String, strExportSheetName As String) As String
    Dim buf As String
    Dim tmp As Variant
    Dim n As Long
    Open strCsvPath For Input As #1
        n = 1
        ' �ŏI�s�܂Ŏ擾
        Do Until EOF(1)
            Line Input #1, buf
            tmp = Split(buf, ",")
            ' �ŏI��܂Ŏ擾����
            For strSearchColumn2 = 1 To UBound(tmp)
                Worksheets(strExportSheetName).Cells(n, strSearchColumn2).Value = tmp(strSearchColumn2)
            Next strSearchColumn2
            n = n + 1
        Loop
    Close #1
    openCsv = "0"
End Function

' �f�[�^���`�̃v���V�W��
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
        ' ���茋�ʕ��J��Ԃ�
        For intSearchCnt = 1 To 12
            Worksheets(strAlignSheetName).Cells(intSearchRow + 1, intSearchCnt).Value = getMeasurementData(strSearchData(intSearchCnt), intSearchRow)
        Next intSearchCnt
    Next intSearchRow
    alignData = "0"
End Function


' ������擾�̃v���V�W��
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


