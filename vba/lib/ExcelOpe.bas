Attribute VB_Name = "ExcelOpe"
Option Explicit

' excel operation library v1.0

Sub CreateSheetList()
    Dim oSheet As Object
    Dim lRowIdx As Long
    Dim lColumnIdx As Long
 
    If MsgBox("�A�N�e�B�u�Z�����牺�ɃV�[�g���ꗗ���쐬���Ă������ł����H", vbYesNo + vbDefaultButton2) = vbNo Then
        'None
    Else
        lRowIdx = ActiveCell.Row
        lColumnIdx = ActiveCell.Column
 
        For Each oSheet In ActiveWorkbook.Sheets
            Cells(lRowIdx, lColumnIdx).Value = oSheet.Name
            lRowIdx = lRowIdx + 1
        Next oSheet
    End If
End Sub

' ============================================
' = �T�v    �A�N�e�B�u�V�[�g�� B2 �Z���ȉ��ɋL�q���ꂽ�֐��������
' =         �֐��c���[���쐬���邽�߂̃I�u�W�F�N�g�𐶐�����B
' =         �֐������L�q���ꂽ�e�L�X�g�{�b�N�X�ƍ��L�ɐڑ����ꂽ�R�l�N�^���A
' =         �֐����̐�������������B
' = �o��    �Ȃ�
' ============================================
Sub CreateFuncTree()
    ' �ϐ���`
    Dim intHeight As Integer                ' �ǉ�����e�L�X�g�{�b�N�X�̈ʒu(����)�
    Dim intStartRow As Integer              ' �X�^�[�g����s��
    Dim intConnecterBeginPointY As Integer  ' �R�l�N�^�n�_�̐����ʒu
    Dim intConnecterBeginPointX As Integer  ' �R�l�N�^�n�_�̐����ʒu
    Dim shpObjectBox As Shape               ' �֐����{�b�N�X��`
    Dim shpObjectLine As Shape              ' �R�l�N�^��`
 
    ' �֐����{�b�N�X�����ʒu��`
    intHeight = 25
    intStartRow = 2
 
    For intSearchRow = intStartRow To (Range("B2").End(xlDown).Row)
        ' === �֐����{�b�N�X���� ===
            ' �I�u�W�F�N�g����
            Set shpObjectBox = ActiveSheet.Shapes.AddShape( _
                Type:=msoShapeFlowchartPredefinedProcess, _
                Left:=200, _
                Top:=intHeight * (intSearchRow - 1), _
                Width:=100, _
                Height:=100)
 
            ' �����ݒ�
            shpObjectBox.Fill.ForeColor.RGB = RGB(128, 0, 0)     ' �w�i�F
            shpObjectBox.Line.ForeColor.RGB = RGB(0, 0, 0)       ' ���̐F
            shpObjectBox.Line.Weight = 2                         ' ���̑���
            shpObjectBox.Select
            Selection.Characters.Text = Cells(intSearchRow, 2)   ' �e�L�X�g�ɐ����̓��e��ݒ�
            Selection.AutoSize = True                            ' �����T�C�Y�����ɂ���
 
        ' === �R�l�N�^���� ===
            ' �I�u�W�F�N�g����
            intConnecterBeginPointY = (intHeight * (intSearchRow - 1)) - 10
            intConnecterBeginPointX = 200
            Set shpObjectLine = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, intConnecterBeginPointX, intConnecterBeginPointY, 0, 0)
            ' �����ݒ�
            shpObjectLine.Line.ForeColor.RGB = RGB(0, 0, 0)      ' ���̐F
            shpObjectLine.Line.Weight = 2                        ' ���̑���
            ' �R�l�N�^�ڑ�
            shpObjectLine.Select
            Selection.ShapeRange.ConnectorFormat.EndConnect shpObjectBox, 1
 
   Next intSearchRow
End Sub

' ==================================================================
' = �T�v    �w�肵���͈͂̕��������������
' =         ��؂蕶�����w�肵���ꍇ�A��������Ԃɕ�����}������
' = ����    rConcRange    Range   [in]  ��������͈�
' = ����    sDlmtr        String  [in]  ��؂蕶��
' = �ߒl                  String        ������̕�����
' = �o��    �Ȃ�
' ==================================================================
Public Function ConcStr( _
    ByRef rConcRange As Range, _
    Optional ByVal sDlmtr As String _
) As Variant
    Dim rConcRangeCnt As Range
    Dim sConcTxtBuf As String
 
    If rConcRange.Rows.Count = 1 Or _
       rConcRange.Columns.Count = 1 Then
        For Each rConcRangeCnt In rConcRange
            sConcTxtBuf = sConcTxtBuf & sDlmtr & rConcRangeCnt.Value
        Next rConcRangeCnt
 
        ' ��؂蕶������
        If sDlmtr <> "" Then
            ConcStr = Mid$(sConcTxtBuf, Len(sDlmtr) + 1)
        Else
            ConcStr = sConcTxtBuf
        End If
    Else
        ConcStr = CVErr(xlErrRef)  '�G���[�l
    End If
End Function
 
' ==================================================================
' = �T�v    ������𕪊����A�w�肵���v�f�̕������ԋp����
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = ����    iExtIndex   String  [in]  ���o����v�f ( 0 origin )
' = �ߒl                String        ���o������
' = �o��    iExtIndex ���v�f�𒴂���ꍇ�A�󕶎����ԋp����
' ==================================================================
Public Function SplitStr( _
    ByVal sStr As String, _
    ByVal sDlmtr As String, _
    ByVal iExtIndex As Integer _
) As String
    Dim vSplitStr As Variant
 
    ' �����񕪊�
    vSplitStr = Split(sStr, sDlmtr)
 
    If iExtIndex > UBound(vSplitStr) Then
        SplitStr = ""
    Else
        SplitStr = vSplitStr(iExtIndex)
    End If
End Function

' ==================================================================
' = �T�v    ���������̗L���𔻒肷�� (TRUE:�L�AFALSE:��)
' = ����    rRange   Range     [in]  �Z��
' = �ߒl             Variant         ���������L��
' = �o��    �Ȃ�
' ==================================================================
Public Function GetStrikeExist( _
    ByRef rRange As Range _
) As Variant
    If rRange.Rows.Count = 1 And _
       rRange.Columns.Count = 1 Then
        GetStrikeExist = rRange.Font.Strikethrough
    Else
        GetStrikeExist = CVErr(xlErrRef)  '�G���[�l
    End If
End Function
 
' ==================================================================
' = �T�v    �t�H���g�J���[��ԋp����
' = ����    rRange   Range     [in]  �Z��
' = �ߒl             Variant         �t�H���g�J���[
' = �o��    �Ȃ�
' ==================================================================
Public Function GetFontColor( _
    ByRef rRange As Range _
) As Variant
    If rRange.Rows.Count = 1 And _
       rRange.Columns.Count = 1 Then
        GetFontColor = rRange.Font.Color
    Else
        GetFontColor = CVErr(xlErrRef)  '�G���[�l
    End If
End Function
 
' ==================================================================
' = �T�v    �w�i�F��ԋp����
' = ����    rRange   Range     [in]  �Z��
' = �ߒl             Variant         �w�i�F
' = �o��    �Ȃ�
' ==================================================================
Public Function GetInteriorColor( _
    ByRef rRange As Range _
) As Variant
    If rRange.Rows.Count = 1 And _
       rRange.Columns.Count = 1 Then
        GetInteriorColor = rRange.Interior.Color
    Else
        GetInteriorColor = CVErr(xlErrRef)  '�G���[�l
    End If
End Function

' ==================================================================
' = �T�v    �r�b�g AND ���Z���s��
' = ����    cInVar1   Currency   [in]  ���͒l �����i10�i�����l�j
' = ����    cInVar2   Currency   [in]  ���͒l �E���i10�i�����l�j
' = �ߒl              Variant          ���Z���ʁi10�i�����l�j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitAnd( _
    ByVal cInVar1 As Currency, _
    ByVal cInVar2 As Currency _
) As Variant
    If (cInVar1 > 2147483647# Or cInVar2 > 2147483647#) Then
        BitAnd = CVErr(xlErrNum)  '�G���[�l
    Else
        BitAnd = cInVar1 And cInVar2
    End If
End Function
 
' ==================================================================
' = �T�v    �r�b�g OR ���Z���s��
' = ����    cInVar1   Currency   [in]  ���͒l �����i10�i�����l�j
' = ����    cInVar2   Currency   [in]  ���͒l �E���i10�i�����l�j
' = �ߒl              Variant          ���Z���ʁi10�i�����l�j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitOr( _
    ByVal cInVar1 As Currency, _
    ByVal cInVar2 As Currency _
) As Variant
    Dim sHexVal As String
    If (cInVar1 > 2147483647# Or cInVar2 > 2147483647#) Then
        BitOr = CVErr(xlErrNum)  '�G���[�l
    Else
        BitOr = cInVar1 Or cInVar2
    End If
End Function
 
' ==================================================================
' = �T�v    �_���V�t�g���s���B
' = ����    cDecVal     Currency  [in]  ���͒l�i10�i�����l�j
' = ����    lShiftNum   Long      [in]  �V�t�g�r�b�g��
' = ����    eDirection  Enum      [in]  �V�t�g�����i�E:0 ��:1�j
' = ����    eShiftType  Enum      [in]  �V�t�g��ʁi�E:�_�� ��:�Z�p�j
' = �ߒl                Variant         �V�t�g���ʁi10�i�����l�j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitShift( _
    ByVal cDecVal As Currency, _
    ByVal lShiftNum As Long, _
    ByVal eDirection As E_SHIFT_DIRECTiON, _
    ByVal eShiftType As E_SHIFT_TYPE _
) As Variant
    Dim sHexVal As String
    Dim cDecValHi As Currency
    Dim cDecValLo As Currency
    Dim sBinVal As String
    Dim cRetVal As Currency
 
    If cDecVal > 4294967295# Or _
       (lShiftNum < 0) Or _
       (eDirection <> RIGHT_SHIFT And eDirection <> LEFT_SHIFT) Then
        BitShift = CVErr(xlErrNum)  '�G���[�l
    Else
        If eShiftType = LOGICAL_SHIFT Then
            'Dec��Hex
            cDecValHi = Int(cDecVal / 2 ^ 16)
            cDecValLo = cDecVal - (cDecValHi * 2 ^ 16)
            sHexVal = UCase(String(4 - Len(Hex(cDecValHi)), "0") & Hex(cDecValHi)) & _
                      UCase(String(4 - Len(Hex(cDecValLo)), "0") & Hex(cDecValLo))
            'Hex��Bin
            sBinVal = Hex2Bin(sHexVal)
            'Shift
            sBinVal = BitLogShiftBin(sBinVal, lShiftNum, eDirection)
            'Bin��Hex
            sHexVal = Bin2Hex(sBinVal)
            'Hex��Dec
            cDecValHi = CCur("&H" & Left$(sHexVal, 4)) * 2 ^ 16
            cDecValLo = CCur("&H" & Right$(sHexVal, 4))
            BitShift = cDecValHi + cDecValLo
        Else
            BitShift = CVErr(xlErrNum) '�Z�p�V�t�g�͔�Ή�
        End If
    End If
End Function
 
'********************************************************************************
'* �����֐���`
'********************************************************************************
Private Function Hex2Bin( _
    ByVal sHexVal As String _
) As String
    Dim sBinVal As String
    Debug.Assert Len(sHexVal) = 8
    Do
        sBinVal = sBinVal & Hex2BinMap(Left$(sHexVal, 1))
        sHexVal = Mid$(sHexVal, 2)
    Loop While sHexVal <> ""
    Hex2Bin = sBinVal
End Function
 
Private Function Bin2Hex( _
    ByVal sBinVal As String _
) As String
    Dim sHexVal As String
    Debug.Assert Len(sBinVal) = 32
    Do
        sHexVal = sHexVal & Bin2HexMap(Left$(sBinVal, 4))
        sBinVal = Mid$(sBinVal, 5)
    Loop While sBinVal <> ""
    Bin2Hex = sHexVal
End Function
 
Private Function BitLogShiftBin( _
    ByVal sBinVal As String, _
    ByVal lShiftNum As Long, _
    ByVal eDirection As E_SHIFT_DIRECTiON _
)
    Debug.Assert Len(sBinVal) = 32
    Debug.Assert lShiftNum >= 0
    If lShiftNum < 32 Then
        Select Case eDirection
            Case RIGHT_SHIFT
                BitLogShiftBin = String(lShiftNum, "0") & Left$(sBinVal, Len(sBinVal) - lShiftNum)
            Case LEFT_SHIFT
                BitLogShiftBin = Right$(sBinVal, Len(sBinVal) - lShiftNum) & String(lShiftNum, "0")
            Case Else
                Debug.Assert False
        End Select
    Else
        BitLogShiftBin = "00000000000000000000000000000000"
    End If
End Function
 
Private Function Hex2BinMap( _
    ByVal sHexVal As String _
) As String
    Select Case sHexVal
        Case "0": Hex2BinMap = "0000"
        Case "1": Hex2BinMap = "0001"
        Case "2": Hex2BinMap = "0010"
        Case "3": Hex2BinMap = "0011"
        Case "4": Hex2BinMap = "0100"
        Case "5": Hex2BinMap = "0101"
        Case "6": Hex2BinMap = "0110"
        Case "7": Hex2BinMap = "0111"
        Case "8": Hex2BinMap = "1000"
        Case "9": Hex2BinMap = "1001"
        Case "A": Hex2BinMap = "1010"
        Case "B": Hex2BinMap = "1011"
        Case "C": Hex2BinMap = "1100"
        Case "D": Hex2BinMap = "1101"
        Case "E": Hex2BinMap = "1110"
        Case "F": Hex2BinMap = "1111"
        Case Else: Debug.Assert False
    End Select
End Function
 
Private Function Bin2HexMap( _
    ByVal sBinVal As String _
) As String
    Select Case sBinVal
        Case "0000": Bin2HexMap = "0"
        Case "0001": Bin2HexMap = "1"
        Case "0010": Bin2HexMap = "2"
        Case "0011": Bin2HexMap = "3"
        Case "0100": Bin2HexMap = "4"
        Case "0101": Bin2HexMap = "5"
        Case "0110": Bin2HexMap = "6"
        Case "0111": Bin2HexMap = "7"
        Case "1000": Bin2HexMap = "8"
        Case "1001": Bin2HexMap = "9"
        Case "1010": Bin2HexMap = "A"
        Case "1011": Bin2HexMap = "B"
        Case "1100": Bin2HexMap = "C"
        Case "1101": Bin2HexMap = "D"
        Case "1110": Bin2HexMap = "E"
        Case "1111": Bin2HexMap = "F"
        Case Else: Debug.Assert False
    End Select
End Function

