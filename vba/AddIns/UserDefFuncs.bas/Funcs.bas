Attribute VB_Name = "Funcs"
Option Explicit

' user define functions v1.0

' ==================================================================
' =  <<�֐��ꗗ>>
' =    �EConcStr            �w�肵���͈͂̕��������������B
' =    �ESplitStr           ������𕪊����A�w�肵���v�f�̕������ԋp����B
' =    �EGetStrNum          �w�蕶����̌���ԋp����B
' =    �ERemoveTailWord     ������؂蕶���ȍ~�̕��������������
' =    �EExtractTailWord    ������؂蕶���ȍ~�̕������ԋp����
' =    �EGetDirPath         �w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����B
' =    �EGetFileName        �w�肳�ꂽ�t�@�C���p�X����t�@�C�����𒊏o����B
' =    �EGetStrikeExist     ���������̗L���𔻒肷��B
' =    �EGetFontColor       �t�H���g�J���[��ԋp����B
' =    �EGetInteriorColor   �w�i�F��ԋp����B
' =    �EBitAnd             �r�b�g AND ���Z���s���B
' =    �EBitOr              �r�b�g OR ���Z���s���B
' =    �EBitShift           �_���V�t�g���s���B
' =    �ERegExpSearch       ���K�\���������s���B
' =    �EConvSnakeToPascal  �����K���ϊ����s���i�X�l�[�N�P�[�X�˃p�X�J���P�[�X�j
' =    �EConvSnakeToCamel   �����K���ϊ����s���i�X�l�[�N�P�[�X�˃L�������P�[�X�j
' =    �EConvCamelToSnake   �����K���ϊ����s���i�L�������P�[�X�˃X�l�[�N�P�[�X�j
' ==================================================================

'********************************************************************************
'* �萔��`
'********************************************************************************
Public Enum E_SHIFT_DIRECTiON
    RIGHT_SHIFT = 0
    LEFT_SHIFT
End Enum
Public Enum E_SHIFT_TYPE
    LOGICAL_SHIFT = 0
    ARITHMETIC_SHIFT '��Ή�
End Enum

'********************************************************************************
'* �O���֐���`
'********************************************************************************
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
' = �T�v    �w�蕶����̌���ԋp����B
' = ����    sTrgtStr      String  [in]  �����Ώە�����
' = ����    sSrchStr      String  [in]  ����������
' = �ߒl                  Long          ������̌�
' = �o��    SplitStr �Ƃ̑g�ݍ��킹�Ńt�@�C�������o�����\�B
' =           ex) B1 = C:\codes\c\Try04.c
' =               B2 = SplitStr( B1, "\", GetStrNum( B2, "\" ) )
' =                 �� Try04.c
' ==================================================================
Public Function GetStrNum( _
    ByVal sTrgtStr As String, _
    ByVal sSrchStr As String _
) As Long
    Dim vSplitStr As Variant
    
    ' �����񕪊�
    vSplitStr = Split(sTrgtStr, sSrchStr)
    
    GetStrNum = UBound(vSplitStr)
End Function

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕������ԋp����B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ���o������
' = �o��    �Ȃ�
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr As String, _
    ByVal sDlmtr As String _
) As String
    Dim asSplitWord() As String
    
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function
Sub Test_ExtractTailWord()
    Debug.Print ExtractTailWord("", "\")               '
    Debug.Print ExtractTailWord("c:\a", "\")           ' a
    Debug.Print ExtractTailWord("c:\a\", "\")          '
    Debug.Print ExtractTailWord("c:\a\b", "\")         ' b
    Debug.Print ExtractTailWord("c:\a\b\", "\")        '
    Debug.Print ExtractTailWord("c:\a\b\c.txt", "\")   ' c.txt
    Debug.Print ExtractTailWord("c:\\b\c.txt", "\")    ' c.txt
    Debug.Print ExtractTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
    Debug.Print ExtractTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
    Debug.Print ExtractTailWord("c:\a\\b\c.txt", "\\") ' b\c.txt
End Sub

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕��������������B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ����������
' = �o��    �Ȃ�
' ==================================================================
Public Function RemoveTailWord( _
    ByVal sStr As String, _
    ByVal sDlmtr As String _
) As String
    Dim sTailWord As String
    Dim lRemoveLen As Long
    
    If sStr = "" Then
        RemoveTailWord = ""
    Else
        If sDlmtr = "" Then
            RemoveTailWord = sStr
        Else
            If InStr(sStr, sDlmtr) = 0 Then
                RemoveTailWord = sStr
            Else
                sTailWord = ExtractTailWord(sStr, sDlmtr)
                lRemoveLen = Len(sDlmtr) + Len(sTailWord)
                RemoveTailWord = Left$(sStr, Len(sStr) - lRemoveLen)
            End If
        End If
    End If
End Function
Private Sub Test_RemoveTailWord()
    Debug.Print RemoveTailWord("", "\") '
    Debug.Print RemoveTailWord("c:\a", "\") '          c:
    Debug.Print RemoveTailWord("c:\a\", "\") '         c:\a
    Debug.Print RemoveTailWord("c:\a\b", "\") '        c:\a
    Debug.Print RemoveTailWord("c:\a\b\", "\") '       c:\a\b
    Debug.Print RemoveTailWord("c:\a\b\c.txt", "\") '  c:\a\b
    Debug.Print RemoveTailWord("c:\\b\c.txt", "\") '   c:\\b
    Debug.Print RemoveTailWord("c:\a\b\c.txt", "") '   c:\a\b\c.txt
    Debug.Print RemoveTailWord("c:\a\b\c.txt", "\\") ' c:\a\b\c.txt
    Debug.Print RemoveTailWord("c:\a\\b\c.txt", "\\") 'c:\a
End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�H���_�p�X
' = �o��    �Ȃ�
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath As String _
) As String
    GetDirPath = RemoveTailWord(sFilePath, "\")
End Function

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�@�C�����𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�@�C����
' = �o��    �Ȃ�
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath As String _
) As String
    GetFileName = ExtractTailWord(sFilePath, "\")
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

' ==================================================================
' = �T�v    ���K�\���������s��
' = ����    sSearchPattern  String   [in]  �����p�^�[��
' = ����    sTargetStr      String   [in]  �����Ώە�����
' = ����    lMatchIdx       Long     [in]  �������ʃC���f�b�N�X�i�����ȗ��j
' = ����    bIsIgnoreCase   Boolean  [in]  ��/��������ʂ��Ȃ����i�����ȗ��j
' = ����    bIsGlobal       Boolean  [in]  ������S�̂��������邩�i�����ȗ��j
' = �ߒl                    Variant        ��������
' = �o��    �Ȃ�
' ==================================================================
Public Function RegExpSearch( _
    ByVal sSearchPattern As String, _
    ByVal sTargetStr As String, _
    Optional ByVal lMatchIdx As Long = 0, _
    Optional ByVal bIsIgnoreCase As Boolean = True, _
    Optional ByVal bIsGlobal As Boolean = True _
) As Variant
    Dim oMatchResult As Object
    Dim oRegExp As Object
    
    Set oRegExp = CreateObject("VBScript.RegExp")
    
    oRegExp.Pattern = sSearchPattern               '�����p�^�[����ݒ�
    oRegExp.IgnoreCase = bIsIgnoreCase             '�啶���Ə���������ʂ��Ȃ�
    oRegExp.Global = bIsGlobal                     '������S�̂�����
    
    Set oMatchResult = oRegExp.Execute(sTargetStr) '�p�^�[���}�b�`���s
    
    If lMatchIdx < 0 Or lMatchIdx > oMatchResult.Count - 1 Then
        RegExpSearch = CVErr(xlErrValue)  '�G���[�l
    Else
        RegExpSearch = oMatchResult(lMatchIdx).Value
    End If
End Function

' ==================================================================
' = �T�v    �����K���ϊ����s���i�X�l�[�N�P�[�X�˃p�X�J���P�[�X�j
' = ����    sInStr      String  [in]  ������i�X�l�[�N�P�[�X�j
' = �ߒl                String        ������i�p�X�J���P�[�X�j
' = �o��    �X�l�[�N�P�[�X                          �c get_input_reader
'           �p�X�J���P�[�X�i�A�b�p�[�L�������P�[�X�j�c GetInputReader
' ==================================================================
Public Function ConvSnakeToPascal( _
    ByVal sInStr As String _
) As String
    sInStr = Replace(sInStr, "_", " ")
    sInStr = StrConv(sInStr, vbProperCase)
    sInStr = Replace(sInStr, " ", "")
    ConvSnakeToPascal = sInStr
End Function

' ==================================================================
' = �T�v    �����K���ϊ����s���i�X�l�[�N�P�[�X�˃L�������P�[�X�j
' = ����    sInStr      String  [in]  ������i�X�l�[�N�P�[�X�j
' = �ߒl                String        ������i�L�������P�[�X�j
' = �o��    �X�l�[�N�P�[�X                          �c get_input_reader
'           �L�������P�[�X�i���[���[�L�������P�[�X�j�c getInputReader
' ==================================================================
Public Function ConvSnakeToCamel( _
    ByVal sInStr As String _
) As String
    If sInStr = "" Then Exit Function
    
    sInStr = Replace(sInStr, "_", " ")
    sInStr = StrConv(sInStr, vbProperCase)
    sInStr = Replace(sInStr, " ", "")
    sInStr = LCase(Left$(sInStr, 1)) & _
             Mid$(sInStr, 2, Len(sInStr))
    ConvSnakeToCamel = sInStr
End Function

' ==================================================================
' = �T�v    �����K���ϊ����s���i�L�������P�[�X�˃X�l�[�N�P�[�X�j
' = ����    sInStr      String  [in]  ������i�X�l�[�N�P�[�X�j
' = �ߒl                String        ������i�L�������P�[�X�j
' = �o��    �L�������P�[�X�i���[���[�L�������P�[�X�j�c getInputReader
'           �X�l�[�N�P�[�X                          �c get_input_reader
' ==================================================================
Public Function ConvCamelToSnake( _
    ByVal sInStr As String _
) As String
    Dim lLoopCnt As Long
    Dim sChar As String
    Dim sRetStr As String
    
    If sInStr = "" Then Exit Function
    
    sRetStr = ""
    For lLoopCnt = 1 To Len(sInStr)
        sChar = Mid$(sInStr, lLoopCnt, 1)
        If sChar = UCase(sChar) Then '�啶��
            sRetStr = sRetStr & "_" & LCase(sChar)
        Else
            sRetStr = sRetStr & sChar
        End If
    Next lLoopCnt
    
    If Left(sRetStr, 1) = "_" Then
        sRetStr = Mid$(sRetStr, 2, Len(sRetStr))
    Else
        'Do Nothing
    End If
    
    ConvCamelToSnake = sRetStr
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
