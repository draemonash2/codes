Attribute VB_Name = "Funcs"
Option Explicit

' user define functions v1.4

' ==================================================================
' =  <<�֐��ꗗ>>
' =    �EConcStr            �w�肵���͈͂̕��������������B
' =    �ESplitStr           ������𕪊����A�w�肵���v�f�̕������ԋp����B
' =    �EGetStrNum          �w�蕶����̌���ԋp����B
' =
' =    �ERemoveTailWord     ������؂蕶���ȍ~�̕��������������
' =    �EExtractTailWord    ������؂蕶���ȍ~�̕������ԋp����
' =    �EGetDirPath         �w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����B
' =    �EGetFileName        �w�肳�ꂽ�t�@�C���p�X����t�@�C�����𒊏o����B
' =
' =    �EGetStrikeExist     ���������̗L���𔻒肷��B
' =    �EGetFontColor       �t�H���g�J���[��ԋp����B
' =    �EGetInteriorColor   �w�i�F��ԋp����B
' =
' =    �EBitAndVal          �r�b�g�`�m�c���Z���s���B�i���l�j
' =    �EBitAndStrHex       �r�b�g�`�m�c���Z���s���B�i������P�U�i���j
' =    �EBitAndStrBin       �r�b�g�`�m�c���Z���s���B�i������Q�i���j
' =    �EBitOrVal           �r�b�g�n�q���Z���s���B�i���l�j
' =    �EBitOrStrHex        �r�b�g�n�q���Z���s���B�i������P�U�i���j
' =    �EBitOrStrBin        �r�b�g�n�q���Z���s���B�i������Q�i���j
' =    �EBitShiftVal        �r�b�g�r�g�h�e�s���Z���s���B�i���l�j
' =    �EBitShiftStrHex     �r�b�g�r�g�h�e�s���Z���s���B�i������P�U�i���j
' =    �EBitShiftStrBin     �r�b�g�r�g�h�e�s���Z���s���B�i������Q�i���j
' =
' =    �ERegExpSearch       ���K�\���������s���B
' =
' =    �EConvSnakeToPascal  �����K���ϊ����s���i�X�l�[�N�P�[�X�˃p�X�J���P�[�X�j
' =    �EConvSnakeToCamel   �����K���ϊ����s���i�X�l�[�N�P�[�X�˃L�������P�[�X�j
' =    �EConvCamelToSnake   �����K���ϊ����s���i�L�������P�[�X�˃X�l�[�N�P�[�X�j
' ==================================================================

'********************************************************************************
'* �萔��`
'********************************************************************************
Public Enum E_SHIFT_DIRECTiON
    LEFT_SHIFT = 0
    RIGHT_SHIFT
End Enum
Public Enum E_SHIFT_TYPE
    LOGICAL_SHIFT = 0
    ARITHMETIC_SHIFT_SIGNBITSAVE
    ARITHMETIC_SHIFT_SIGNBITTRUNC
End Enum
Public Enum E_SHIFT_ARIGN
    ARIGN_EIGHTBIT = 0
    '  �o�͌��ʂ�8�r�b�g���E�ɑ�����B
    '    ex1) 10101011 ���E1�r�b�g�V�t�g
    '      �� 11010101
    '    ex2) 10101011 ����1�r�b�g�V�t�g
    '      �� 1111111101010110
    ARIGN_NO
    '  �o�͌��ʂ�8�r�b�g���E�ɑ����Ȃ��B
    '    ex1) 10101011 ���E1�r�b�g�V�t�g
    '      ��  1010101
    '    ex2) 10101011 ����1�r�b�g�V�t�g
    '      �� 101010110
End Enum

'********************************************************************************
'* �O���֐���`
'********************************************************************************
' ==================================================================
' = �T�v    �w�肵���͈͂̕��������������
' =         ��؂蕶�����w�肵���ꍇ�A��������Ԃɕ�����}������
' = ����    rConcRange    Range   [in]  ��������͈�
' = ����    sDlmtr        String  [in]  ��؂蕶��
' = �ߒl                  Variant       ������̕�����
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
    Private Sub Test_ConcStr()
        'Range�^��VBA������͂ł��Ȃ����߁A�e�X�g�ł��Ȃ��B
    End Sub

' ==================================================================
' = �T�v    ������𕪊����A�w�肵���v�f�̕������ԋp����
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = ����    iExtIndex   String  [in]  ���o����v�f ( 0 origin )
' = �ߒl                Variant       ���o������
' = �o��    iExtIndex ���v�f�𒴂���ꍇ�A�󕶎����ԋp����
' ==================================================================
Public Function SplitStr( _
    ByVal sStr As String, _
    ByVal sDlmtr As String, _
    ByVal iExtIndex As Integer _
) As Variant
    If sDlmtr = "" Then
        SplitStr = sStr
    Else
        If sStr = "" Then
            SplitStr = ""
        Else
            Dim vSplitStr As Variant
            vSplitStr = Split(sStr, sDlmtr) ' �����񕪊�
            If iExtIndex > UBound(vSplitStr) Or _
               iExtIndex < LBound(vSplitStr) Then
                SplitStr = ""
            Else
                SplitStr = vSplitStr(iExtIndex)
            End If
        End If
    End If
End Function
    Private Sub Test_SplitStr()
        Debug.Print "*** test start! ***"
        Debug.Print SplitStr("c:\test\a.txt", "\", 0)  'c:
        Debug.Print SplitStr("c:\test\a.txt", "\", 1)  'test
        Debug.Print SplitStr("c:\test\a.txt", "\", 2)  'a.txt
        Debug.Print SplitStr("c:\test\a.txt", "\", -1) '
        Debug.Print SplitStr("c:\test\a.txt", "\", 3)  '
        Debug.Print SplitStr("", "\", 1)               '
        Debug.Print SplitStr("c:\a.txt", "", 1)        'c:\a.txt
        Debug.Print SplitStr("", "", 1)                '
        Debug.Print SplitStr("", "", 0)                '
        Debug.Print "*** test finished! ***"
    End Sub

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
    
    If sTrgtStr = "" Then
        GetStrNum = 0
    Else
        ' �����񕪊�
        vSplitStr = Split(sTrgtStr, sSrchStr)
        GetStrNum = UBound(vSplitStr)
    End If
End Function
    Private Sub Test_GetStrNum()
        Debug.Print "*** test start! ***"
        Debug.Print GetStrNum("", "\")                '0
        Debug.Print GetStrNum("c:\a.txt", "\")        '1
        Debug.Print GetStrNum("c:\test\a.txt", "\")   '2
        Debug.Print GetStrNum("c:\test\a.txt", "")    '0
        Debug.Print GetStrNum("c:\test\\a.txt", "\")  '3
        Debug.Print GetStrNum("c:\test\\a.txt", "\\") '1
        Debug.Print "*** test finished! ***"
    End Sub

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
    Private Sub Test_ExtractTailWord()
        Debug.Print "*** test start! ***"
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
        Debug.Print "*** test finished! ***"
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
        Debug.Print "*** test start! ***"
        Debug.Print RemoveTailWord("", "\")               '
        Debug.Print RemoveTailWord("c:\a", "\")           ' c:
        Debug.Print RemoveTailWord("c:\a\", "\")          ' c:\a
        Debug.Print RemoveTailWord("c:\a\b", "\")         ' c:\a
        Debug.Print RemoveTailWord("c:\a\b\", "\")        ' c:\a\b
        Debug.Print RemoveTailWord("c:\a\b\c.txt", "\")   ' c:\a\b
        Debug.Print RemoveTailWord("c:\\b\c.txt", "\")    ' c:\\b
        Debug.Print RemoveTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
        Debug.Print RemoveTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
        Debug.Print RemoveTailWord("c:\a\\b\c.txt", "\\") ' c:\a
        Debug.Print "*** test finished! ***"
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
    Private Sub Test_GetDirPath()
        'RemoveTailWord�Ɠ����̃e�X�g�P�[�X�̂��߁A�e�X�g���Ȃ�
    End Sub

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
    Private Sub Test_GetFileName()
        'ExtractTailWord�Ɠ����̃e�X�g�P�[�X�̂��߁A�e�X�g���Ȃ�
    End Sub

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
    Private Sub Test_GetStrikeExist()
        'Range�^��VBA������͂ł��Ȃ����߁A�e�X�g�ł��Ȃ��B
    End Sub

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
    Private Sub Test_GetFontColor()
        'Range�^��VBA������͂ł��Ȃ����߁A�e�X�g�ł��Ȃ��B
    End Sub

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
    Private Sub Test_GetInteriorColor()
        'Range�^��VBA������͂ł��Ȃ����߁A�e�X�g�ł��Ȃ��B
    End Sub

' ==================================================================
' = �T�v    �r�b�g�`�m�c���Z���s���B�i���l�j
' = ����    cInVal1   Currency   [in]  ���͒l �����i10�i�����l�j
' = ����    cInVal2   Currency   [in]  ���͒l �E���i10�i�����l�j
' = �ߒl              Variant          ���Z���ʁi10�i�����l�j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitAndVal( _
    ByVal cInVal1 As Currency, _
    ByVal cInVal2 As Currency _
) As Variant
    If (cInVal1 > 2147483647# Or cInVal2 > 2147483647#) Then
        BitAndVal = CVErr(xlErrNum)  '�G���[�l
    Else
        BitAndVal = cInVal1 And cInVal2
    End If
End Function
    Private Sub Test_BitAndVal()
        Debug.Print "*** test start! ***"
        Debug.Print Hex(BitAndVal(&HFFFF&, &HFF00&)) 'FF00
        Debug.Print Hex(BitAndVal(&HFFFF&, &HFF&))   'FF
        Debug.Print Hex(BitAndVal(&HFFFF&, &HA5A5&)) 'A5A5
        Debug.Print Hex(BitAndVal(&HA5&, &HA500&))   '0
        Debug.Print Hex(BitAndVal(&H1&, &H8&))       '0
        Debug.Print Hex(BitAndVal(&H1&, &HA&))       '0
        Debug.Print Hex(BitAndVal(&H5&, &HA&))       '0
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�`�m�c���Z���s���B�i������P�U�i���j
' = ����    sInHexVal1  String     [in]  ���͒l �����i������j
' = ����    sInHexVal2  String     [in]  ���͒l �E���i������j
' = ����    lInDigitNum Long       [in]  �o�͌���
' = �ߒl                Variant          ���Z���ʁi������j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitAndStrHex( _
    ByVal sInHexVal1 As String, _
    ByVal sInHexVal2 As String, _
    Optional ByVal lInDigitNum As Long = 0 _
) As Variant
    If sInHexVal1 = "" Or sInHexVal2 = "" Then
        BitAndStrHex = CVErr(xlErrNull) '�G���[�l
        Exit Function
    End If
    If lInDigitNum < 0 Then
        BitAndStrHex = CVErr(xlErrNum) '�G���[�l
        Exit Function
    End If
    
    '���͒l�P�̂a�h�m�ϊ�
    Dim sInBinVal1 As String
    sInBinVal1 = Hex2Bin(sInHexVal1)
    If sInBinVal1 = "error" Then
        BitAndStrHex = CVErr(xlErrValue) '�G���[�l
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal1 : " & sInBinVal1
    
    '���͒l�Q�̂a�h�m�ϊ�
    Dim sInBinVal2 As String
    sInBinVal2 = Hex2Bin(sInHexVal2)
    If sInBinVal2 = "error" Then
        BitAndStrHex = CVErr(xlErrValue) '�G���[�l
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal2 : " & sInBinVal2
    
    '�a�h�m �`�m�c���Z
    Dim sOutBinVal As String
    sOutBinVal = BitAndStrBin(sInBinVal1, sInBinVal2, lInDigitNum * 4)
    
    '�a�h�m�˂g�d�w�ϊ�
    Dim sOutHexVal As String
    sOutHexVal = Bin2Hex(sOutBinVal, True)
    Debug.Assert sOutHexVal <> "error"
    BitAndStrHex = sOutHexVal
    
End Function
    Private Sub Test_BitAndStrHex()
        Debug.Print "*** test start! ***"
        Debug.Print BitAndStrHex("FF", "FF00")                  '0000
        Debug.Print BitAndStrHex("A5A5", "5A5A")                '0000
        Debug.Print BitAndStrHex("A5A5", "00FF")                '00A5
        Debug.Print BitAndStrHex("A5", "00FF")                  '00A5
        Debug.Print BitAndStrHex("FFFF0B00", "01010300")        '01010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 10)    '0001010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 8)     '01010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 7)     '1010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 6)     '010300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 5)     '10300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 4)     '0300
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 2)     '00
        Debug.Print BitAndStrHex("FFFF0B00", "01010300", 1)     '0
        Debug.Print BitAndStrHex("ab", "00FF")                  '00AB
        Debug.Print BitAndStrHex("cd", "00FF")                  '00CD
        Debug.Print BitAndStrHex("ef", "00FF")                  '00EF
        Debug.Print BitAndStrHex(" 0B00", "0300")               '�G���[ 2015
        Debug.Print BitAndStrHex("", "0300")                    '�G���[ 2000
        Debug.Print BitAndStrHex("0B00", "")                    '�G���[ 2000
        Debug.Print BitAndStrHex("0B00", "0300", -1)            '�G���[ 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�`�m�c���Z���s���B�i������Q�i���j
' = ����    sInBinVal1  String     [in]  ���͒l �����i������j
' = ����    sInBinVal2  String     [in]  ���͒l �E���i������j
' = ����    lInBitLen   Long       [in]  �o�̓r�b�g��
' = �ߒl                Variant          ���Z���ʁi������j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitAndStrBin( _
    ByVal sInBinVal1 As String, _
    ByVal sInBinVal2 As String, _
    Optional ByVal lInBitLen As Long = 0 _
) As Variant
    If sInBinVal1 = "" Or sInBinVal2 = "" Then
        BitAndStrBin = CVErr(xlErrNum) '�G���[�l
    Else
        '�����������킹��
        If Len(sInBinVal1) > Len(sInBinVal2) Then
            sInBinVal2 = String(Len(sInBinVal1) - Len(sInBinVal2), "0") & sInBinVal2
        Else
            sInBinVal1 = String(Len(sInBinVal2) - Len(sInBinVal1), "0") & sInBinVal1
        End If
        Debug.Assert Len(sInBinVal1) = Len(sInBinVal2)
        
        'OR���Z
        Dim lValIdx As Long
        Dim sOutBin As String
        Dim bIsError As Boolean
        lValIdx = Len(sInBinVal1)
        sOutBin = ""
        bIsError = False
        Do
            Select Case Mid$(sInBinVal1, lValIdx, 1) & Mid$(sInBinVal2, lValIdx, 1)
                Case "00": sOutBin = "0" & sOutBin
                Case "10": sOutBin = "0" & sOutBin
                Case "01": sOutBin = "0" & sOutBin
                Case "11": sOutBin = "1" & sOutBin
                Case Else: bIsError = True
            End Select
            lValIdx = lValIdx - 1
        Loop While lValIdx > 0 And bIsError = False
        
        If bIsError = True Then
            BitAndStrBin = CVErr(xlErrNum) '�G���[�l
        Else
            If lInBitLen = 0 Then
                BitAndStrBin = sOutBin
            Else
                If lInBitLen <= Len(sOutBin) Then
                    BitAndStrBin = Right$(sOutBin, lInBitLen)
                Else
                    BitAndStrBin = String(lInBitLen - Len(sOutBin), "0") & sOutBin
                End If
            End If
        End If
    End If
End Function
    Private Sub Test_BitAndStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print BitAndStrBin("11", "11110")          '00010
        Debug.Print BitAndStrBin("11110", "11")          '00010
        Debug.Print BitAndStrBin("1", "1")               '1
        Debug.Print BitAndStrBin("1", "0")               '0
        Debug.Print BitAndStrBin("0", "1")               '0
        Debug.Print BitAndStrBin("0", "0")               '0
        Debug.Print BitAndStrBin("00000011", "11000000") '00000000
        Debug.Print BitAndStrBin("0111", "0010", 10)     '0000000010
        Debug.Print BitAndStrBin("0111", "0010", 0)      '0010
        Debug.Print BitAndStrBin("0111", "0010", 2)      '10
        Debug.Print BitAndStrBin("0111", "0010", 1)      '0
        Debug.Print BitAndStrBin("0101", "001F")         '�G���[ 2036
        Debug.Print BitAndStrBin(" 101", "0010")         '�G���[ 2036
        Debug.Print BitAndStrBin("", "0010")             '�G���[ 2036
        Debug.Print BitAndStrBin("0101", "")             '�G���[ 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�n�q���Z���s���B�i���l�j
' = ����    cInVal1   Currency   [in]  ���͒l �����i10�i�����l�j
' = ����    cInVal2   Currency   [in]  ���͒l �E���i10�i�����l�j
' = �ߒl              Variant          ���Z���ʁi10�i�����l�j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitOrVal( _
    ByVal cInVal1 As Currency, _
    ByVal cInVal2 As Currency _
) As Variant
    Dim sHexVal As String
    If (cInVal1 > 2147483647# Or cInVal2 > 2147483647#) Then
        BitOrVal = CVErr(xlErrNum)  '�G���[�l
    Else
        BitOrVal = cInVal1 Or cInVal2
    End If
End Function
    Private Sub Test_BitOrVal()
        Debug.Print "*** test start! ***"
        Debug.Print Hex(BitOrVal(&HFFFF&, &HFF00&)) 'FFFF
        Debug.Print Hex(BitOrVal(&HFFFF&, &HFF&))   'FFFF
        Debug.Print Hex(BitOrVal(&HFFFF&, &HA5A5&)) 'FFFF
        Debug.Print Hex(BitOrVal(&HA5&, &HA500&))   'A5A5
        Debug.Print Hex(BitOrVal(&H1&, &H8&))       '9
        Debug.Print Hex(BitOrVal(&H1&, &HA&))       'B
        Debug.Print Hex(BitOrVal(&H5&, &HA&))       'F
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�n�q���Z���s���B�i������P�U�i���j
' = ����    sInHexVal1  String     [in]  ���͒l �����i������j
' = ����    sInHexVal2  String     [in]  ���͒l �E���i������j
' = ����    lInDigitNum Long       [in]  �o�͌���
' = �ߒl                Variant          ���Z���ʁi������j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitOrStrHex( _
    ByVal sInHexVal1 As String, _
    ByVal sInHexVal2 As String, _
    Optional ByVal lInDigitNum As Long = 0 _
) As Variant
    If sInHexVal1 = "" Or sInHexVal2 = "" Then
        BitOrStrHex = CVErr(xlErrNull) '�G���[�l
        Exit Function
    End If
    If lInDigitNum < 0 Then
        BitOrStrHex = CVErr(xlErrNum) '�G���[�l
        Exit Function
    End If
    
    '���͒l�P�̂a�h�m�ϊ�
    Dim sInBinVal1 As String
    sInBinVal1 = Hex2Bin(sInHexVal1)
    If sInBinVal1 = "error" Then
        BitOrStrHex = CVErr(xlErrValue) '�G���[�l
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal1 : " & sInBinVal1
    
    '���͒l�Q�̂a�h�m�ϊ�
    Dim sInBinVal2 As String
    sInBinVal2 = Hex2Bin(sInHexVal2)
    If sInBinVal2 = "error" Then
        BitOrStrHex = CVErr(xlErrValue) '�G���[�l
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal2 : " & sInBinVal2
    
    '�a�h�m �n�q���Z
    Dim sOutBinVal As String
    sOutBinVal = BitOrStrBin(sInBinVal1, sInBinVal2, lInDigitNum * 4)
    
    '�a�h�m�˂g�d�w�ϊ�
    Dim sOutHexVal As String
    sOutHexVal = Bin2Hex(sOutBinVal, True)
    Debug.Assert sOutHexVal <> "error"
    BitOrStrHex = sOutHexVal
End Function
    Private Sub Test_BitOrStrHex()
        Debug.Print "*** test start! ***"
        Debug.Print BitOrStrHex("FF", "FF00")                  'FFFF
        Debug.Print BitOrStrHex("A5A5", "5A5A")                'FFFF
        Debug.Print BitOrStrHex("A5A5", "00FF")                'A5FF
        Debug.Print BitOrStrHex("A5", "00FF")                  '00FF
        Debug.Print BitOrStrHex("FFFF0800", "01010300")        'FFFF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 10)    '00FFFF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 8)     'FFFF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 7)     'FFF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 6)     'FF0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 5)     'F0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 4)     '0B00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 2)     '00
        Debug.Print BitOrStrHex("FFFF0800", "01010300", 1)     '0
        Debug.Print BitOrStrHex("ab", "0000")                  '00AB
        Debug.Print BitOrStrHex("cd", "0000")                  '00CD
        Debug.Print BitOrStrHex("ef", "0000")                  '00EF
        Debug.Print BitOrStrHex(" 0800", "0300")               '�G���[ 2015
        Debug.Print BitOrStrHex("", "0300")                    '�G���[ 2000
        Debug.Print BitOrStrHex("0800", "")                    '�G���[ 2000
        Debug.Print BitOrStrHex("0800", "0300", -1)            '�G���[ 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�n�q���Z���s���B�i������Q�i���j
' = ����    sInBinVal1  String     [in]  ���͒l �����i������j
' = ����    sInBinVal2  String     [in]  ���͒l �E���i������j
' = ����    lInBitLen   Long       [in]  �o�̓r�b�g��
' = �ߒl                Variant          ���Z���ʁi������j
' = �o��    �Ȃ�
' ==================================================================
Public Function BitOrStrBin( _
    ByVal sInBinVal1 As String, _
    ByVal sInBinVal2 As String, _
    Optional ByVal lInBitLen As Long = 0 _
) As Variant
    If sInBinVal1 = "" Or sInBinVal2 = "" Then
        BitOrStrBin = CVErr(xlErrNum) '�G���[�l
    Else
        '�����������킹��
        If Len(sInBinVal1) > Len(sInBinVal2) Then
            sInBinVal2 = String(Len(sInBinVal1) - Len(sInBinVal2), "0") & sInBinVal2
        Else
            sInBinVal1 = String(Len(sInBinVal2) - Len(sInBinVal1), "0") & sInBinVal1
        End If
        Debug.Assert Len(sInBinVal1) = Len(sInBinVal2)
        
        'OR���Z
        Dim lValIdx As Long
        Dim sOutBin As String
        Dim bIsError As Boolean
        lValIdx = Len(sInBinVal1)
        sOutBin = ""
        bIsError = False
        Do
            Select Case Mid$(sInBinVal1, lValIdx, 1) & Mid$(sInBinVal2, lValIdx, 1)
                Case "00": sOutBin = "0" & sOutBin
                Case "10": sOutBin = "1" & sOutBin
                Case "01": sOutBin = "1" & sOutBin
                Case "11": sOutBin = "1" & sOutBin
                Case Else: bIsError = True
            End Select
            lValIdx = lValIdx - 1
        Loop While lValIdx > 0 And bIsError = False
        
        If bIsError = True Then
            BitOrStrBin = CVErr(xlErrNum) '�G���[�l
        Else
            If lInBitLen = 0 Then
                BitOrStrBin = sOutBin
            Else
                If lInBitLen <= Len(sOutBin) Then
                    BitOrStrBin = Right$(sOutBin, lInBitLen)
                Else
                    BitOrStrBin = String(lInBitLen - Len(sOutBin), "0") & sOutBin
                End If
            End If
        End If
    End If
End Function
    Private Sub Test_BitOrStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print BitOrStrBin("11", "11110")                  '11111
        Debug.Print BitOrStrBin("11110", "11")                  '11111
        Debug.Print BitOrStrBin("1", "1")                       '1
        Debug.Print BitOrStrBin("1", "0")                       '1
        Debug.Print BitOrStrBin("0", "1")                       '1
        Debug.Print BitOrStrBin("0", "0")                       '0
        Debug.Print BitOrStrBin("00000011", "11000000")         '11000011
        Debug.Print BitOrStrBin("01010101", "00010010", 0)      '01010111
        Debug.Print BitOrStrBin("01010101", "00010010", 10)     '0001010111
        Debug.Print BitOrStrBin("01010101", "00010010", 7)      '1010111
        Debug.Print BitOrStrBin("01010101", "00010010", 6)      '010111
        Debug.Print BitOrStrBin("01010101", "00010010", 4)      '0111
        Debug.Print BitOrStrBin("01010101", "00010010", 3)      '111
        Debug.Print BitOrStrBin("01010101", "00010010", 2)      '11
        Debug.Print BitOrStrBin("01010101", "00010010", 1)      '1
        Debug.Print BitOrStrBin("0101", "001F")                 '�G���[ 2036
        Debug.Print BitOrStrBin(" 101", "0010")                 '�G���[ 2036
        Debug.Print BitOrStrBin("K01", "0010")                  '�G���[ 2036
        Debug.Print BitOrStrBin("", "0010")                     '�G���[ 2036
        Debug.Print BitOrStrBin("0101", "")                     '�G���[ 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�r�g�h�e�s���Z���s���B�i���l�j
' = ����    cInDecVal       Currency  [in]  ���͒l�i10�i�����l�j
' = ����    lInShiftNum     Long      [in]  �V�t�g�r�b�g��
' = ����    eInDirection    Enum      [in]  �V�t�g�����i0:�� 1:�E�j
' = ����    eInShiftType    Enum      [in]  �V�t�g��ʁi0:�_�� 1:�Z�p�j
' = �ߒl                    Variant         �V�t�g���ʁi10�i�����l�j
' = �o��    ���Z�p�V�t�g�͔�Ή���
' ==================================================================
Public Function BitShiftVal( _
    ByVal cInDecVal As Currency, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    Optional ByVal eInShiftType As E_SHIFT_TYPE = LOGICAL_SHIFT _
) As Variant
    If cInDecVal > 4294967295# Then
        BitShiftVal = CVErr(xlErrNum)  '�G���[�l
        Exit Function
    End If
    If lInShiftNum < 0 Then
        BitShiftVal = CVErr(xlErrNum)  '�G���[�l
        Exit Function
    End If
    If eInDirection <> RIGHT_SHIFT And eInDirection <> LEFT_SHIFT Then
        BitShiftVal = CVErr(xlErrValue)  '�G���[�l
        Exit Function
    End If
    If eInShiftType <> LOGICAL_SHIFT And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITSAVE And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITTRUNC Then
        BitShiftVal = CVErr(xlErrValue)  '�G���[�l
        Exit Function
    End If
    
    Dim sHexVal As String
    Dim cInDecValHi As Currency
    Dim cInDecValLo As Currency
    Dim sBinVal As String
    Dim cRetVal As Currency
    If eInShiftType = LOGICAL_SHIFT Then
        On Error Resume Next
        'Dec��Hex
        cInDecValHi = Int(cInDecVal / 2 ^ 16)
        cInDecValLo = cInDecVal - (cInDecValHi * 2 ^ 16)
        sHexVal = UCase(String(4 - Len(Hex(cInDecValHi)), "0") & Hex(cInDecValHi)) & _
                  UCase(String(4 - Len(Hex(cInDecValLo)), "0") & Hex(cInDecValLo))
        'Hex��Bin
        sBinVal = Hex2Bin(sHexVal)
        'Shift
        sBinVal = BitShiftStrBin(sBinVal, lInShiftNum, eInDirection, eInShiftType, 32)
        'Bin��Hex
        sHexVal = Bin2Hex(sBinVal, True)
        'Hex��Dec
        cInDecValHi = CCur("&H" & Left$(sHexVal, 4)) * 2 ^ 16
        cInDecValLo = CCur("&H" & Right$(sHexVal, 4))
        cRetVal = cInDecValHi + cInDecValLo
        If Err.Number <> 0 Then
            BitShiftVal = CVErr(xlErrNum) '�G���[�l
            Err.Clear
        Else
            BitShiftVal = cRetVal
        End If
        On Error GoTo 0
    Else
        BitShiftVal = CVErr(xlErrNA) '���Z�p�V�t�g�͔�Ή���
    End If
End Function
    Private Sub Test_BitShiftVal()
        Debug.Print "*** test start! ***"
        Debug.Print Hex(BitShiftVal(&H10&, 0, RIGHT_SHIFT, LOGICAL_SHIFT))      '10
        Debug.Print Hex(BitShiftVal(&H10&, 1, RIGHT_SHIFT, LOGICAL_SHIFT))      '8
        Debug.Print Hex(BitShiftVal(&H10&, 2, RIGHT_SHIFT, LOGICAL_SHIFT))      '4
        Debug.Print Hex(BitShiftVal(&H10&, 3, RIGHT_SHIFT, LOGICAL_SHIFT))      '2
        Debug.Print Hex(BitShiftVal(&H10&, 4, RIGHT_SHIFT, LOGICAL_SHIFT))      '1
        Debug.Print Hex(BitShiftVal(&H10&, 5, RIGHT_SHIFT, LOGICAL_SHIFT))      '0
        Debug.Print Hex(BitShiftVal(&H10&, 8, RIGHT_SHIFT, LOGICAL_SHIFT))      '0
        Debug.Print Hex(BitShiftVal(&H10&, 0, LEFT_SHIFT, LOGICAL_SHIFT))       '10
        Debug.Print Hex(BitShiftVal(&H10&, 1, LEFT_SHIFT, LOGICAL_SHIFT))       '20
        Debug.Print Hex(BitShiftVal(&H10&, 2, LEFT_SHIFT, LOGICAL_SHIFT))       '40
        Debug.Print Hex(BitShiftVal(&H10&, 3, LEFT_SHIFT, LOGICAL_SHIFT))       '80
        Debug.Print Hex(BitShiftVal(&H10&, 8, LEFT_SHIFT, LOGICAL_SHIFT))       '1000
        Debug.Print Hex(BitShiftVal(&H10&, 12, LEFT_SHIFT, LOGICAL_SHIFT))      '10000
        Debug.Print Hex(BitShiftVal(&H10&, 16, LEFT_SHIFT, LOGICAL_SHIFT))      '100000
        Debug.Print Hex(BitShiftVal(&H10&, 20, LEFT_SHIFT, LOGICAL_SHIFT))      '1000000
        Debug.Print Hex(BitShiftVal(&H10&, 24, LEFT_SHIFT, LOGICAL_SHIFT))      '10000000
        Debug.Print Hex(BitShiftVal(&H10&, 25, LEFT_SHIFT, LOGICAL_SHIFT))      '20000000
        Debug.Print Hex(BitShiftVal(&H10&, 26, LEFT_SHIFT, LOGICAL_SHIFT))      '40000000
       'Debug.Print Hex(BitShiftVal(&H10&, 27, LEFT_SHIFT, LOGICAL_SHIFT))      '�G���[�iHex()�ɂăI�[�o�[�t���[�j
        Debug.Print BitShiftVal(&H10&, 27, LEFT_SHIFT, LOGICAL_SHIFT)           '2147483648
        Debug.Print BitShiftVal(&H10&, 28, LEFT_SHIFT, LOGICAL_SHIFT)           '0
        Debug.Print BitShiftVal(&H10&, 29, LEFT_SHIFT, LOGICAL_SHIFT)           '0
        Debug.Print Hex(BitShiftVal(&H7FFFFFFF, 0, LEFT_SHIFT, LOGICAL_SHIFT))  '7FFFFFFF
       'Debug.Print BitShiftVal(&H80000000, 0, LEFT_SHIFT, LOGICAL_SHIFT)       '80000000���o�O�H
       'Debug.Print BitShiftVal(4294967294#, 0, LEFT_SHIFT, LOGICAL_SHIFT)      'FFFFFFFF���o�O�H
        Debug.Print BitShiftVal(&H10&, 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE) '�G���[ 2042
        Debug.Print BitShiftVal(&H10&, -1, LEFT_SHIFT, LOGICAL_SHIFT)           '�G���[ 2036
        Debug.Print BitShiftVal(&H10&, 1, 3, LOGICAL_SHIFT)                     '�G���[ 2015
        Debug.Print BitShiftVal(&H10&, 1, LEFT_SHIFT, 3)                        '�G���[ 2015
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�r�g�h�e�s���Z���s���B�i������P�U�i���j
' = ����    sInHexVal       String  [in]    ���͒l�i������j
' = ����    lInShiftNum     Long    [in]    �V�t�g�r�b�g��
' = ����    eInDirection    Enum    [in]    �V�t�g�����i0:�� 1:�E�j
' = ����    eInShiftType    Enum    [in]    �V�t�g���
' =                                           0:�_��
' =                                           1:�Z�p�i�����r�b�g�ێ��j(��1)
' =                                           2:�Z�p�i�����r�b�g�؎́j(��1)
' = ����    lInDigitNum     Long    [in]    �o�͌���
' = �ߒl                    Variant         �V�t�g���ʁi������j
' = �o��    (��1) �o�͌��������͒l�i������j�̒��������������ꍇ�ɁA
' =               �����r�b�g��ێ����邩�A�������Đ؂�̂Ă邩��I������B
' =           ex1) 10101011 ���o�͌���4�Ƃ��ĉE1�Z�p(�����r�b�g�ێ�)�V�t�g �� 1101
' =           ex2) 10101011 ���o�͌���4�Ƃ��ĉE1�Z�p(�����r�b�g�؎�)�V�t�g �� 0101
' ==================================================================
Public Function BitShiftStrHex( _
    ByVal sInHexVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    Optional ByVal eInShiftType As E_SHIFT_TYPE = LOGICAL_SHIFT, _
    Optional ByVal lInDigitNum As Long = 0 _
) As Variant
    If sInHexVal = "" Then
        BitShiftStrHex = CVErr(xlErrNull) '�G���[�l
        Exit Function
    End If
    If lInShiftNum < 0 Then
        BitShiftStrHex = CVErr(xlErrNum) '�G���[�l
        Exit Function
    End If
    If eInShiftType <> LOGICAL_SHIFT And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITSAVE And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITTRUNC Then
        BitShiftStrHex = CVErr(xlErrValue) '�G���[�l
        Exit Function
    End If
    If eInDirection <> LEFT_SHIFT And eInDirection <> RIGHT_SHIFT Then
        BitShiftStrHex = CVErr(xlErrValue) '�G���[�l
        Exit Function
    End If
    If lInDigitNum < 0 Then
        BitShiftStrHex = CVErr(xlErrNum) '�G���[�l
        Exit Function
    End If
    
    '���͒l�̂g�d�w�˂a�h�m�ϊ�
    Dim sInBinVal As String
    sInBinVal = Hex2Bin(sInHexVal)
    If sInBinVal = "error" Then
        BitShiftStrHex = CVErr(xlErrValue) '�G���[�l
        Exit Function
    Else
        'Do Nothing
    End If
   'Debug.Print "sInBinVal : " & sInBinVal
    
    '�a�h�m�V�t�g
    Dim sOutBinVal As String
    Dim sTmpBinVal As String
    sTmpBinVal = BitShiftStrBin(sInBinVal, lInShiftNum, eInDirection, eInShiftType, lInDigitNum * 4)
    Dim lModNum As Long
    lModNum = Len(sTmpBinVal) Mod 4
    If lModNum = 0 Then
        sOutBinVal = sTmpBinVal
    Else
        sOutBinVal = String(4 - lModNum, "0") & sTmpBinVal
    End If
    
    '�a�h�m�˂g�d�w�ϊ�
    Dim sOutHexVal As String
    sOutHexVal = Bin2Hex(sOutBinVal, True)
    Debug.Assert sOutHexVal <> "error"
    BitShiftStrHex = sOutHexVal
End Function
    Private Sub Test_BitShiftStrHex()
        Debug.Print "*** test start! ***"
        Debug.Print BitShiftStrHex("B0", 0, LEFT_SHIFT, LOGICAL_SHIFT)                  'B0
        Debug.Print BitShiftStrHex("B0", 1, LEFT_SHIFT, LOGICAL_SHIFT)                  '160
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT)                  '2C0
        Debug.Print BitShiftStrHex("B0", 3, LEFT_SHIFT, LOGICAL_SHIFT)                  '580
        Debug.Print BitShiftStrHex("B0", 4, LEFT_SHIFT, LOGICAL_SHIFT)                  'B00
        Debug.Print BitShiftStrHex("B0", 120, LEFT_SHIFT, LOGICAL_SHIFT)                'B0 + 0�~30��
        Debug.Print BitShiftStrHex("B0", 0, RIGHT_SHIFT, LOGICAL_SHIFT)                 'B0
        Debug.Print BitShiftStrHex("B0", 1, RIGHT_SHIFT, LOGICAL_SHIFT)                 '58
        Debug.Print BitShiftStrHex("B0", 2, RIGHT_SHIFT, LOGICAL_SHIFT)                 '2C
        Debug.Print BitShiftStrHex("B0", 3, RIGHT_SHIFT, LOGICAL_SHIFT)                 '16
        Debug.Print BitShiftStrHex("B0", 4, RIGHT_SHIFT, LOGICAL_SHIFT)                 'B
        Debug.Print BitShiftStrHex("B0", 7, RIGHT_SHIFT, LOGICAL_SHIFT)                 '1
        Debug.Print BitShiftStrHex("B0", 8, RIGHT_SHIFT, LOGICAL_SHIFT)                 '0
        Debug.Print BitShiftStrHex("B0", 9, RIGHT_SHIFT, LOGICAL_SHIFT)                 '0
        Debug.Print BitShiftStrHex("B0", 120, RIGHT_SHIFT, LOGICAL_SHIFT)               '0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 0)               '2C0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 1)               '0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 2)               'C0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 3)               '2C0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 4)               '02C0
        Debug.Print BitShiftStrHex("B0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 10)              '00000002C0
        Debug.Print BitShiftStrHex("B0", 9, RIGHT_SHIFT, LOGICAL_SHIFT, 8)              '00000000
        Debug.Print BitShiftStrHex("B0", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE)   '160
        Debug.Print BitShiftStrHex("ab", 8, LEFT_SHIFT)                                 'AB00
        Debug.Print BitShiftStrHex("cd", 8, LEFT_SHIFT)                                 'CD00
        Debug.Print BitShiftStrHex("ef", 8, LEFT_SHIFT)                                 'EF00
        Debug.Print BitShiftStrHex("", 2, LEFT_SHIFT, LOGICAL_SHIFT)                    '�G���[ 2000
        Debug.Print BitShiftStrHex(" B", 2, RIGHT_SHIFT, LOGICAL_SHIFT)                 '�G���[ 2015
        Debug.Print BitShiftStrHex("K", 2, RIGHT_SHIFT, LOGICAL_SHIFT)                  '�G���[ 2015
        Debug.Print BitShiftStrHex("B", -1, LEFT_SHIFT, LOGICAL_SHIFT)                  '�G���[ 2036
        Debug.Print BitShiftStrHex("B", 1, 3, LOGICAL_SHIFT)                            '�G���[ 2015
        Debug.Print BitShiftStrHex("B", 1, LEFT_SHIFT, 3)                               '�G���[ 2015
        Debug.Print BitShiftStrHex("B", 1, LEFT_SHIFT, , -1)                            '�G���[ 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�r�g�h�e�s���Z���s���B�i������Q�i���j
' = ����    sInBinVal       String  [in]    ���͒l�i������j
' = ����    lInShiftNum     Long    [in]    �V�t�g�r�b�g��
' = ����    eInDirection    Enum    [in]    �V�t�g�����i0:�� 1:�E�j
' = ����    eInShiftType    Enum    [in]    �V�t�g���
' =                                           0:�_��
' =                                           1:�Z�p�i�����r�b�g�ێ��j(��1)
' =                                           2:�Z�p�i�����r�b�g�؎́j(��1)
' = ����    lInBitLen       Long    [in]    �o�̓r�b�g��
' = �ߒl                    Variant         �V�t�g���ʁi������j
' = �o��    (��1) �o�͌��������͒l�i������j�̒��������������ꍇ�ɁA
' =               �����r�b�g��ێ����邩�A�������Đ؂�̂Ă邩��I������B
' =           ex1) "AB"(0b10101011) ���o�͌���1�Ƃ��ĉE1�Z�p(�����r�b�g�ێ�)�V�t�g
' =              �� "D"(0b1101)
' =           ex2) "AB"(0b10101011) ���o�͌���1�Ƃ��ĉE1�Z�p(�����r�b�g�؎�)�V�t�g
' =              �� "5"(0b0101)
' ==================================================================
Public Function BitShiftStrBin( _
    ByVal sInBinVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    Optional ByVal eInShiftType As E_SHIFT_TYPE = LOGICAL_SHIFT, _
    Optional ByVal lInBitLen As Long = 0 _
) As Variant
    If sInBinVal = "" Then
        BitShiftStrBin = CVErr(xlErrNull) '�G���[�l
        Exit Function
    End If
    If lInShiftNum < 0 Then
        BitShiftStrBin = CVErr(xlErrNum) '�G���[�l
        Exit Function
    End If
    If Replace(Replace(sInBinVal, "1", ""), "0", "") <> "" Then
        BitShiftStrBin = CVErr(xlErrValue) '�G���[�l
        Exit Function
    End If
    If eInShiftType <> LOGICAL_SHIFT And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITSAVE And _
       eInShiftType <> ARITHMETIC_SHIFT_SIGNBITTRUNC Then
        BitShiftStrBin = CVErr(xlErrValue) '�G���[�l
        Exit Function
    End If
    If eInDirection <> LEFT_SHIFT And eInDirection <> RIGHT_SHIFT Then
        BitShiftStrBin = CVErr(xlErrValue) '�G���[�l
        Exit Function
    End If
    If lInBitLen < 0 Then
        BitShiftStrBin = CVErr(xlErrNum) '�G���[�l
        Exit Function
    End If
    
    Select Case eInShiftType
        Case LOGICAL_SHIFT:
            BitShiftStrBin = BitShiftLogStrBin(sInBinVal, lInShiftNum, eInDirection, eInShiftType, lInBitLen)
        Case ARITHMETIC_SHIFT_SIGNBITSAVE:
            'BitShiftStrBin = CVErr(xlErrNA) '�Z�p�V�t�g�͖�����
            BitShiftStrBin = BitShiftAriStrBin(sInBinVal, lInShiftNum, eInDirection, eInShiftType, lInBitLen)
        Case ARITHMETIC_SHIFT_SIGNBITTRUNC:
            BitShiftStrBin = BitShiftAriStrBin(sInBinVal, lInShiftNum, eInDirection, eInShiftType, lInBitLen)
        Case Else
            BitShiftStrBin = CVErr(xlErrValue) '�G���[�l
    End Select
End Function
    Private Sub Test_BitShiftStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print BitShiftStrBin("1011", 0, LEFT_SHIFT, LOGICAL_SHIFT)                        '1011
        Debug.Print BitShiftStrBin("1011", 1, LEFT_SHIFT, LOGICAL_SHIFT)                        '10110
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT)                        '101100
        Debug.Print BitShiftStrBin("1011", 4, LEFT_SHIFT, LOGICAL_SHIFT)                        '10110000
        Debug.Print BitShiftStrBin("1011", 0, RIGHT_SHIFT, LOGICAL_SHIFT)                       '1011
        Debug.Print BitShiftStrBin("1011", 1, RIGHT_SHIFT, LOGICAL_SHIFT)                       '101
        Debug.Print BitShiftStrBin("1011", 2, RIGHT_SHIFT, LOGICAL_SHIFT)                       '10
        Debug.Print BitShiftStrBin("1011", 3, RIGHT_SHIFT, LOGICAL_SHIFT)                       '1
        Debug.Print BitShiftStrBin("1011", 4, RIGHT_SHIFT, LOGICAL_SHIFT)                       '0
        Debug.Print BitShiftStrBin("1011", 5, RIGHT_SHIFT, LOGICAL_SHIFT)                       '0
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 0)                     '101100
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 2)                     '00
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 3)                     '100
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 4)                     '1100
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 10)                    '0000101100
        Debug.Print BitShiftStrBin("1011", 120, LEFT_SHIFT, LOGICAL_SHIFT)                      '1011 + 0�~120��
        Debug.Print BitShiftStrBin("1011", 5, RIGHT_SHIFT, LOGICAL_SHIFT, 8)                    '00000000
        Debug.Print BitShiftStrBin("10001011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 16) '1111111000101100
        Debug.Print BitShiftStrBin("", 2, LEFT_SHIFT, LOGICAL_SHIFT)                            '�G���[ 2000
        Debug.Print BitShiftStrBin(":1011", 2, LEFT_SHIFT, LOGICAL_SHIFT)                       '�G���[ 2015
        Debug.Print BitShiftStrBin("1021", 2, LEFT_SHIFT, LOGICAL_SHIFT)                        '�G���[ 2015
        Debug.Print BitShiftStrBin("1011", -1, LEFT_SHIFT, LOGICAL_SHIFT)                       '�G���[ 2036
        Debug.Print BitShiftStrBin("1011", 2, 3, LOGICAL_SHIFT)                                 '�G���[ 2015
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, 3)                                    '�G���[ 2015
        Debug.Print BitShiftStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, -1)                    '�G���[ 2036
        Debug.Print "*** test finished! ***"
    End Sub

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
    ByVal sTargetStr As String, _
    ByVal sSearchPattern As String, _
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
    Private Sub Test_RegExpSearch()
        Dim sTargetStr As String
        sTargetStr = "void TestFunc(int arg1, char arg2);"
        Debug.Print "*** test start! ***"
        Debug.Print RegExpSearch(sTargetStr, " \w+\(")
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �����K���ϊ����s���i�X�l�[�N�P�[�X�˃p�X�J���P�[�X�j
' = ����    sInStr      String  [in]  ������i�X�l�[�N�P�[�X�j
' = �ߒl                Variant       ������i�p�X�J���P�[�X�j
' = �o��    �E�X�l�[�N�P�[�X�ƃp�X�J���P�[�X�ɂ���
' =             - �X�l�[�N�P�[�X �c get_input_reader
' =             - �p�X�J���P�[�X �c GetInputReader
' =               �i���A�b�p�[�L�������P�[�X�j
' =         �EsInStr �͒P��݂̂��w�肷�邱�ƁB�i�󔒂��܂߂Ȃ��j
' ==================================================================
Public Function ConvSnakeToPascal( _
    ByVal sInStr As String _
) As Variant
    If InStr(sInStr, " ") > 0 Then
        ConvSnakeToPascal = CVErr(xlErrName) '�G���[�l
    Else
        sInStr = Replace(sInStr, "_", " ")
        sInStr = StrConv(sInStr, vbProperCase)
        sInStr = Replace(sInStr, " ", "")
        ConvSnakeToPascal = sInStr
    End If
End Function
    Private Sub Test_ConvSnakeToPascal()
        Debug.Print "*** test start! ***"
        Debug.Print ConvSnakeToPascal("get_input_reader") 'GetInputReader
        Debug.Print ConvSnakeToPascal("getinputreader")   'Getinputreader
        Debug.Print ConvSnakeToPascal("GetInputReader")   'GetInputReader
        Debug.Print ConvSnakeToPascal("get input reader") '�G���[ 2029
        Debug.Print ConvSnakeToPascal("")                 '
        Debug.Print ConvSnakeToPascal("get_")             'Get
        Debug.Print ConvSnakeToPascal("_get_")            'Get
        Debug.Print ConvSnakeToPascal("get input_reader") '�G���[ 2029
        Debug.Print ConvSnakeToPascal("get-input-reader") 'Get-input-reader
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �����K���ϊ����s���i�X�l�[�N�P�[�X�˃L�������P�[�X�j
' = ����    sInStr      String  [in]  ������i�X�l�[�N�P�[�X�j
' = �ߒl                Variant       ������i�L�������P�[�X�j
' = �o��    �E�X�l�[�N�P�[�X�ƃL�������P�[�X�ɂ���
' =             - �X�l�[�N�P�[�X �c get_input_reader
' =             - �L�������P�[�X �c getInputReader
' =               �i�����[���[�L�������P�[�X�j
' =         �EsInStr �͒P��݂̂��w�肷�邱�ƁB�i�󔒂��܂߂Ȃ��j
' ==================================================================
Public Function ConvSnakeToCamel( _
    ByVal sInStr As String _
) As Variant
    If InStr(sInStr, " ") > 0 Then
        ConvSnakeToCamel = CVErr(xlErrName) '�G���[�l
    Else
        sInStr = Replace(sInStr, "_", " ")
        sInStr = StrConv(sInStr, vbProperCase)
        sInStr = Replace(sInStr, " ", "")
        sInStr = LCase(Left$(sInStr, 1)) & _
                 Mid$(sInStr, 2, Len(sInStr))
        ConvSnakeToCamel = sInStr
    End If
End Function
    Private Sub Test_ConvSnakeToCamel()
        Debug.Print "*** test start! ***"
        Debug.Print ConvSnakeToCamel("get_input_reader") 'getInputReader
        Debug.Print ConvSnakeToCamel("getinputreader")   'getinputreader
        Debug.Print ConvSnakeToCamel("GetInputReader")   'getinputreader
        Debug.Print ConvSnakeToCamel("get input reader") '�G���[ 2029
        Debug.Print ConvSnakeToCamel("")                 '
        Debug.Print ConvSnakeToCamel("get_")             'get
        Debug.Print ConvSnakeToCamel("_get_")            'get
        Debug.Print ConvSnakeToCamel("get input_reader") '�G���[ 2029
        Debug.Print ConvSnakeToCamel("get-input-reader") 'get-input-reader
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �����K���ϊ����s���i�L�������P�[�X�˃X�l�[�N�P�[�X�j
' = ����    sInStr      String  [in]  ������i�X�l�[�N�P�[�X�j
' = �ߒl                Variant       ������i�L�������P�[�X�j
' = �o��    �E�L�������P�[�X�ƃX�l�[�N�P�[�X�ɂ���
' =             - �L�������P�[�X �c getInputReader
' =               �i�����[���[�L�������P�[�X�j
' =             - �X�l�[�N�P�[�X �c get_input_reader
' =         �EsInStr �͒P��݂̂��w�肷�邱�ƁB�i�󔒂��܂߂Ȃ��j
' =         �E�A�b�p�[�L�������P�[�X���w��\�B
' ==================================================================
Public Function ConvCamelToSnake( _
    ByVal sInStr As String _
) As Variant
    If InStr(sInStr, " ") > 0 Then
        ConvCamelToSnake = CVErr(xlErrName) '�G���[�l
    Else
        If sInStr = "" Then
            ConvCamelToSnake = ""
        Else
            Dim lLoopCnt As Long
            Dim sChar As String
            Dim sRetStr As String
            
            sRetStr = ""
            For lLoopCnt = 1 To Len(sInStr)
                sChar = Mid$(sInStr, lLoopCnt, 1)
                If sChar <> "_" Then             '
                    If sChar = UCase(sChar) Then '�啶��
                        If lLoopCnt = 1 Then     '�ꕶ����
                            sRetStr = sRetStr & LCase(sChar)
                        Else
                            sRetStr = sRetStr & "_" & LCase(sChar)
                        End If
                    Else
                        sRetStr = sRetStr & sChar
                    End If
                Else
                    sRetStr = sRetStr & sChar
                End If
            Next lLoopCnt
            
            ConvCamelToSnake = sRetStr
        End If
    End If
End Function
    Private Sub Test_ConvCamelToSnake()
        Debug.Print "*** test start! ***"
        Debug.Print ConvCamelToSnake("getInputReader")  'get_input_reader
        Debug.Print ConvCamelToSnake("GetInputReader")  'get_input_reader
        Debug.Print ConvCamelToSnake("getInput_Reader") 'get_input__reader
        Debug.Print ConvCamelToSnake("GInput_Reader")   'g_input__reader
        Debug.Print ConvCamelToSnake("GINPUT")          'g_i_n_p_u_t
        Debug.Print ConvCamelToSnake("")                '
        Debug.Print "*** test finished! ***"
    End Sub

'********************************************************************************
'* �����֐���`
'********************************************************************************

'�w��͈͈ȊO�̒l���w�肷��ƕ����� "error" ��ԋp����B
Private Function Hex2Bin( _
    ByVal sHexVal As String _
) As String
    Dim sOutBinVal As String
    Dim sTmpBinVal As String
    Dim lIdx As Long
    Dim sChar As String
    If sHexVal = "" Then
        sOutBinVal = ""
    Else
        sOutBinVal = ""
        For lIdx = 1 To Len(sHexVal)
            sChar = Mid$(sHexVal, lIdx, 1)
            sTmpBinVal = Hex2BinMap(sChar)
            If sTmpBinVal = "error" Then
                sOutBinVal = sTmpBinVal
                Exit For
            Else
                sOutBinVal = sOutBinVal & sTmpBinVal
            End If
        Next lIdx
    End If
    Hex2Bin = sOutBinVal
End Function
    Private Sub Test_Hex2Bin()
        Debug.Print "*** test start! ***"
        Debug.Print Hex2Bin("0123")      '0000000100100011
        Debug.Print Hex2Bin("4567")      '0100010101100111
        Debug.Print Hex2Bin("89AB")      '1000100110101011
        Debug.Print Hex2Bin("CDEF")      '1100110111101111
        Debug.Print Hex2Bin("cdef")      '1100110111101111
        Debug.Print Hex2Bin("c")         '1100
        Debug.Print Hex2Bin("01234567")  '00000001001000110100010101100111
        Debug.Print Hex2Bin("")          '
        Debug.Print Hex2Bin("ab ")       'error
        Debug.Print Hex2Bin(":cd")       'error
        Debug.Print "*** test finished! ***"
    End Sub

'�w��͈͈ȊO�̒l���w�肷��ƕ����� "error" ��ԋp����B
Private Function Bin2Hex( _
    ByVal sBinVal As String, _
    ByVal bIsUcase As Boolean _
) As String
    Dim sExtBinStr As String
    Dim sTmpHexVal As String
    Dim sOutHexVal As String
    Dim lIdx As Long
    If sBinVal = "" Then
        sOutHexVal = ""
    Else
        If Len(sBinVal) Mod 4 = 0 Then
            For lIdx = 1 To Len(sBinVal) Step 4
                sExtBinStr = Mid$(sBinVal, lIdx, 4)
                sTmpHexVal = Bin2HexMap(sExtBinStr, bIsUcase)
                If sTmpHexVal = "error" Then
                    sOutHexVal = "error"
                    Exit For
                Else
                    sOutHexVal = sOutHexVal & sTmpHexVal
                End If
            Next lIdx
        Else
            sOutHexVal = "error"
        End If
    End If
    Bin2Hex = sOutHexVal
End Function
    Private Sub Test_Bin2Hex()
        Debug.Print "*** test start! ***"
        Debug.Print Bin2Hex("0000000100100011", True)                   '0123
        Debug.Print Bin2Hex("0100010101100111", True)                   '4567
        Debug.Print Bin2Hex("1000100110101011", True)                   '89AB
        Debug.Print Bin2Hex("1100110111101111", True)                   'CDEF
        Debug.Print Bin2Hex("1100110111101111", False)                  'cdef
        Debug.Print Bin2Hex("1100", False)                              'c
        Debug.Print Bin2Hex("00000001001000110100010101100111", True)   '01234567
        Debug.Print Bin2Hex("", True)                                   '
        Debug.Print Bin2Hex("110011011110111", False)                   'error
        Debug.Print Bin2Hex("010 ", True)                               'error
        Debug.Print Bin2Hex(":011", True)                               'error
        Debug.Print "*** test finished! ***"
    End Sub

'�w��͈͈ȊO�̒l���w�肷��ƕ����� "error" ��ԋp����B
Private Function Hex2BinMap( _
    ByVal sHexVal As String _
) As String
    Select Case UCase(sHexVal)
        Case "0":  Hex2BinMap = "0000"
        Case "1":  Hex2BinMap = "0001"
        Case "2":  Hex2BinMap = "0010"
        Case "3":  Hex2BinMap = "0011"
        Case "4":  Hex2BinMap = "0100"
        Case "5":  Hex2BinMap = "0101"
        Case "6":  Hex2BinMap = "0110"
        Case "7":  Hex2BinMap = "0111"
        Case "8":  Hex2BinMap = "1000"
        Case "9":  Hex2BinMap = "1001"
        Case "A":  Hex2BinMap = "1010"
        Case "B":  Hex2BinMap = "1011"
        Case "C":  Hex2BinMap = "1100"
        Case "D":  Hex2BinMap = "1101"
        Case "E":  Hex2BinMap = "1110"
        Case "F":  Hex2BinMap = "1111"
        Case Else: Hex2BinMap = "error"
    End Select
End Function

'�w��͈͈ȊO�̒l���w�肷��ƕ����� "error" ��ԋp����B
Private Function Bin2HexMap( _
    ByVal sBinVal As String, _
    ByVal bIsUcase As Boolean _
) As String
    If bIsUcase = True Then
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
            Case Else:   Bin2HexMap = "error"
        End Select
    Else
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
            Case "1010": Bin2HexMap = "a"
            Case "1011": Bin2HexMap = "b"
            Case "1100": Bin2HexMap = "c"
            Case "1101": Bin2HexMap = "d"
            Case "1110": Bin2HexMap = "e"
            Case "1111": Bin2HexMap = "f"
            Case Else:   Bin2HexMap = "error"
        End Select
    End If
End Function

'�_���r�b�g�V�t�g�i������Łj
Private Function BitShiftLogStrBin( _
    ByVal sInBinVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    ByVal eInShiftType As E_SHIFT_TYPE, _
    ByVal lInBitLen As Long _
) As String
    Debug.Assert sInBinVal <> ""
    Debug.Assert Replace(Replace(sInBinVal, "1", ""), "0", "") = ""
    Debug.Assert lInShiftNum >= 0
    Debug.Assert eInDirection = LEFT_SHIFT Or RIGHT_SHIFT
    Debug.Assert eInShiftType = LOGICAL_SHIFT
    Debug.Assert lInBitLen >= 0
    
    '�r�b�g�V�t�g
    Dim sTmpBinVal As String
    Select Case eInDirection
        Case RIGHT_SHIFT
            If Len(sInBinVal) > lInShiftNum Then
                sTmpBinVal = Left$(sInBinVal, Len(sInBinVal) - lInShiftNum)
            Else
                sTmpBinVal = "0"
            End If
        Case LEFT_SHIFT
            sTmpBinVal = sInBinVal & String(lInShiftNum, "0")
        Case Else
            Debug.Assert 0
    End Select
    
    '�r�b�g�ʒu���킹
    If lInBitLen = 0 Then
        BitShiftLogStrBin = sTmpBinVal
    Else
        If lInBitLen > Len(sTmpBinVal) Then
            BitShiftLogStrBin = String(lInBitLen - Len(sTmpBinVal), "0") & sTmpBinVal
        ElseIf lInBitLen < Len(sTmpBinVal) Then
            BitShiftLogStrBin = Right$(sTmpBinVal, lInBitLen)
        Else
            BitShiftLogStrBin = sTmpBinVal
        End If
    End If
End Function
    Private Sub Test_BitShiftLogStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print BitShiftLogStrBin("0", 0, LEFT_SHIFT, LOGICAL_SHIFT, 8)      '00000000
        Debug.Print BitShiftLogStrBin("0", 2, LEFT_SHIFT, LOGICAL_SHIFT, 8)      '00000000
        Debug.Print BitShiftLogStrBin("1", 0, LEFT_SHIFT, LOGICAL_SHIFT, 8)      '00000001
        Debug.Print BitShiftLogStrBin("1", 2, LEFT_SHIFT, LOGICAL_SHIFT, 8)      '00000100
        Debug.Print BitShiftLogStrBin("1", 7, LEFT_SHIFT, LOGICAL_SHIFT, 8)      '10000000
        Debug.Print BitShiftLogStrBin("1", 8, LEFT_SHIFT, LOGICAL_SHIFT, 8)      '00000000
        Debug.Print BitShiftLogStrBin("1", 0, RIGHT_SHIFT, LOGICAL_SHIFT, 8)     '00000001
        Debug.Print BitShiftLogStrBin("1", 1, RIGHT_SHIFT, LOGICAL_SHIFT, 8)     '00000000
        Debug.Print BitShiftLogStrBin("1", 2, RIGHT_SHIFT, LOGICAL_SHIFT, 8)     '00000000
        Debug.Print BitShiftLogStrBin("1011", 0, LEFT_SHIFT, LOGICAL_SHIFT, 0)   '1011
        Debug.Print BitShiftLogStrBin("1011", 1, LEFT_SHIFT, LOGICAL_SHIFT, 0)   '10110
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 0)   '101100
        Debug.Print BitShiftLogStrBin("1011", 0, RIGHT_SHIFT, LOGICAL_SHIFT, 0)  '1011
        Debug.Print BitShiftLogStrBin("1011", 1, RIGHT_SHIFT, LOGICAL_SHIFT, 0)  '101
        Debug.Print BitShiftLogStrBin("1011", 2, RIGHT_SHIFT, LOGICAL_SHIFT, 0)  '10
        Debug.Print BitShiftLogStrBin("1011", 3, RIGHT_SHIFT, LOGICAL_SHIFT, 0)  '1
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, LOGICAL_SHIFT, 0)  '0
        Debug.Print BitShiftLogStrBin("1011", 5, RIGHT_SHIFT, LOGICAL_SHIFT, 0)  '0
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 2)   '00
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 3)   '100
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 4)   '1100
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, LOGICAL_SHIFT, 10)  '0000101100
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, LOGICAL_SHIFT, 8)  '00000000
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, LOGICAL_SHIFT, 8)  '00000000
       'Debug.Print BitShiftLogStrBin("", 2, LEFT_SHIFT, LOGICAL_SHIFT, 10)       '�v���O������~
       'Debug.Print BitShiftLogStrBin("101A", 1, LEFT_SHIFT, LOGICAL_SHIFT, 10)   '�v���O������~
       'Debug.Print BitShiftLogStrBin("1011", -1, LEFT_SHIFT, LOGICAL_SHIFT, 10)  '�v���O������~
       'Debug.Print BitShiftLogStrBin("1011", 1, 5, LOGICAL_SHIFT, 10)            '�v���O������~
       'Debug.Print BitShiftLogStrBin("0", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '�v���O������~
       'Debug.Print BitShiftLogStrBin("0", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 8) '�v���O������~
       'Debug.Print BitShiftLogStrBin("0", 0, LEFT_SHIFT, 3, 8)                   '�v���O������~
        Debug.Print "*** test finished! ***"
    End Sub

'�Z�p�r�b�g�V�t�g�i������Łj
Private Function BitShiftAriStrBin( _
    ByVal sInBinVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    ByVal eInShiftType As E_SHIFT_TYPE, _
    ByVal lInBitLen As Long _
) As String
    '<<E_SHIFT_ARIGN>>
    '  ARIGN_EIGHTBIT : �o�͌��ʂ�8�r�b�g���E�ɑ�����
    '                     ex1) 10101011 ���E1�r�b�g�V�t�g
    '                       �� 11010101
    '                     ex2) 10101011 ����1�r�b�g�V�t�g
    '                       �� 1111111101010110
    '  ARIGN_NO       : �o�͌��ʂ�8�r�b�g���E�ɑ����Ȃ�
    '                     ex1) 10101011 ���E1�r�b�g�V�t�g
    '                       ��  1010101
    '                     ex2) 10101011 ����1�r�b�g�V�t�g
    '                       �� 101010110
    Dim eArignType As E_SHIFT_ARIGN
    eArignType = ARIGN_NO
    'eArignType = ARIGN_EIGHTBIT
    
    Debug.Assert sInBinVal <> ""
    Debug.Assert Replace(Replace(sInBinVal, "1", ""), "0", "") = ""
    Debug.Assert lInShiftNum >= 0
    Debug.Assert eInDirection = LEFT_SHIFT Or RIGHT_SHIFT
    Debug.Assert eInShiftType = ARITHMETIC_SHIFT_SIGNBITSAVE Or ARITHMETIC_SHIFT_SIGNBITTRUNC
    Debug.Assert lInBitLen >= 0
    If eArignType = ARIGN_EIGHTBIT Then
        Debug.Assert Len(sInBinVal) = 8
        Debug.Assert lInBitLen Mod 8 = 0
    Else
        'Do Nothing
    End If
    
    '�V�t�g
    Dim sTmpBinVal As String
    Dim sOutLogicBit As String
    Dim sInSignBit As String
    Dim sInLogicBit As String
    sInSignBit = Left$(sInBinVal, 1)
    sInLogicBit = Mid$(sInBinVal, 2, Len(sInBinVal))
    Select Case eInDirection
        Case RIGHT_SHIFT
            If Len(sInLogicBit) > lInShiftNum Then
                sOutLogicBit = Left$(sInLogicBit, Len(sInLogicBit) - lInShiftNum)
            Else
                sOutLogicBit = ""
            End If
        Case LEFT_SHIFT
            sOutLogicBit = sInLogicBit & String(lInShiftNum, "0")
        Case Else
            Debug.Assert 0
    End Select
    sTmpBinVal = sInSignBit & sOutLogicBit
    
    '�r�b�g�ʒu���킹
    Dim sPadBit As String
    If lInBitLen = 0 Then
        If eArignType = ARIGN_EIGHTBIT Then
            If ((Len(sOutLogicBit) + 1) Mod 8) = 0 Then
                sPadBit = ""
            Else
                sPadBit = String(8 - ((Len(sOutLogicBit) + 1) Mod 8), sInSignBit)
            End If
            BitShiftAriStrBin = sPadBit & sInSignBit & sOutLogicBit
        Else
            BitShiftAriStrBin = sTmpBinVal
        End If
    Else
        If lInBitLen > Len(sTmpBinVal) Then
            BitShiftAriStrBin = String(lInBitLen - Len(sTmpBinVal), sInSignBit) & sTmpBinVal
        ElseIf lInBitLen < Len(sTmpBinVal) Then
            Select Case eInShiftType
                Case ARITHMETIC_SHIFT_SIGNBITSAVE:
                    BitShiftAriStrBin = sInSignBit & Right$(sOutLogicBit, lInBitLen - 1)
                Case ARITHMETIC_SHIFT_SIGNBITTRUNC:
                    BitShiftAriStrBin = Right$(sTmpBinVal, lInBitLen)
                Case LOGICAL_SHIFT:
                    Debug.Assert 0
                Case Else:
                    Debug.Assert 0
            End Select
        Else
            BitShiftAriStrBin = sTmpBinVal
        End If
    End If
End Function
    Private Sub Test_BitShiftAriStrBin()
        Dim eArignType As E_SHIFT_ARIGN
        eArignType = ARIGN_NO
        'eArignType = ARIGN_EIGHTBIT
        
        Debug.Print "*** test start! ***"
        If eArignType = ARIGN_NO Then
            Debug.Print "<<test 01-01>>"
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '01001011
            Debug.Print BitShiftAriStrBin("01001011", 1, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '0100101
            Debug.Print BitShiftAriStrBin("01001011", 2, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '010010
            Debug.Print BitShiftAriStrBin("01001011", 3, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '01001
            Debug.Print BitShiftAriStrBin("01001011", 4, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '0100
            Debug.Print BitShiftAriStrBin("01001011", 5, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '010
            Debug.Print BitShiftAriStrBin("01001011", 6, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '01
            Debug.Print BitShiftAriStrBin("01001011", 7, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '0
            Debug.Print BitShiftAriStrBin("01001011", 8, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '0
            Debug.Print BitShiftAriStrBin("01001011", 9, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '0
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10) '0001001011
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 5)  '01011
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 4)  '0011
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 3)  '011
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 2)  '01
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 1)  '0
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '01001011
            Debug.Print BitShiftAriStrBin("01001011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '010010110
            Debug.Print BitShiftAriStrBin("01001011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '0100101100
            Debug.Print BitShiftAriStrBin("01001011", 5, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '0100101100000
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10)  '0001001011
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 5)   '01011
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 4)   '0011
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 3)   '011
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 2)   '01
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 1)   '0
            Debug.Print "<<test 01-02>>"
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '01001011
            Debug.Print BitShiftAriStrBin("01001011", 1, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '0100101
            Debug.Print BitShiftAriStrBin("01001011", 2, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '010010
            Debug.Print BitShiftAriStrBin("01001011", 3, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '01001
            Debug.Print BitShiftAriStrBin("01001011", 4, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '0100
            Debug.Print BitShiftAriStrBin("01001011", 5, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '010
            Debug.Print BitShiftAriStrBin("01001011", 6, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '01
            Debug.Print BitShiftAriStrBin("01001011", 7, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '0
            Debug.Print BitShiftAriStrBin("01001011", 8, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '0
            Debug.Print BitShiftAriStrBin("01001011", 9, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '0
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 10) '0001001011
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 5)  '01011
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 4)  '1011
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 3)  '011
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 2)  '11
            Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 1)  '1
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)   '01001011
            Debug.Print BitShiftAriStrBin("01001011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)   '010010110
            Debug.Print BitShiftAriStrBin("01001011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)   '0100101100
            Debug.Print BitShiftAriStrBin("01001011", 5, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)   '0100101100000
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 10)  '0001001011
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 5)   '01011
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 4)   '1011
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 3)   '011
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 2)   '11
            Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 1)   '1
            Debug.Print "<<test 02-01>>"
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '10001011
            Debug.Print BitShiftAriStrBin("10001011", 1, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '1000101
            Debug.Print BitShiftAriStrBin("10001011", 2, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '100010
            Debug.Print BitShiftAriStrBin("10001011", 3, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '10001
            Debug.Print BitShiftAriStrBin("10001011", 4, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '1000
            Debug.Print BitShiftAriStrBin("10001011", 5, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '100
            Debug.Print BitShiftAriStrBin("10001011", 6, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '10
            Debug.Print BitShiftAriStrBin("10001011", 7, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '1
            Debug.Print BitShiftAriStrBin("10001011", 8, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '1
            Debug.Print BitShiftAriStrBin("10001011", 9, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '1
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10) '1110001011
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 5)  '11011
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 4)  '1011
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 3)  '111
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 2)  '11
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 1)  '1
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '10001011
            Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '100010110
            Debug.Print BitShiftAriStrBin("10001011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '1000101100
            Debug.Print BitShiftAriStrBin("10001011", 5, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '1000101100000
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10)  '1110001011
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 5)   '11011
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 4)   '1011
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 3)   '111
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 2)   '11
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 1)   '1
            Debug.Print "<<test 02-02>>"
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '10001011
            Debug.Print BitShiftAriStrBin("10001011", 1, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '1000101
            Debug.Print BitShiftAriStrBin("10001011", 2, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '100010
            Debug.Print BitShiftAriStrBin("10001011", 3, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '10001
            Debug.Print BitShiftAriStrBin("10001011", 4, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '1000
            Debug.Print BitShiftAriStrBin("10001011", 5, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '100
            Debug.Print BitShiftAriStrBin("10001011", 6, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '10
            Debug.Print BitShiftAriStrBin("10001011", 7, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '1
            Debug.Print BitShiftAriStrBin("10001011", 8, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '1
            Debug.Print BitShiftAriStrBin("10001011", 9, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)  '1
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 10) '1110001011
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 5)  '01011
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 4)  '1011
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 3)  '011
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 2)  '11
            Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 1)  '1
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)   '10001011
            Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)   '100010110
            Debug.Print BitShiftAriStrBin("10001011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)   '1000101100
            Debug.Print BitShiftAriStrBin("10001011", 5, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 0)   '1000101100000
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 10)  '1110001011
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 5)   '01011
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 4)   '1011
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 3)   '011
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 2)   '11
            Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITTRUNC, 1)   '1
        '   Debug.Print "<<test 03>>"
        '   Debug.Print BitShiftAriStrBin("10001011", 1, 5, ARITHMETIC_SHIFT_SIGNBITSAVE, 10)           '�v���O������~
        '   Debug.Print BitShiftAriStrBin("", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10)          '�v���O������~
        '   Debug.Print BitShiftAriStrBin("10001011", -1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10) '�v���O������~
        '   Debug.Print BitShiftAriStrBin("1000101A", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10)  '�v���O������~
        '   Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, -1)  '�v���O������~
        Else
            Debug.Print "<<test 01>>"
            Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '00101011
            Debug.Print BitShiftAriStrBin("00101011", 1, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '00010101
            Debug.Print BitShiftAriStrBin("00101011", 2, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '00001010
            Debug.Print BitShiftAriStrBin("00101011", 3, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '00000101
            Debug.Print BitShiftAriStrBin("00101011", 4, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '00000010
            Debug.Print BitShiftAriStrBin("00101011", 5, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '00000001
            Debug.Print BitShiftAriStrBin("00101011", 6, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '00000000
            Debug.Print BitShiftAriStrBin("00101011", 7, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '00000000
            Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 16) '0000000000101011
            Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '00101011
            Debug.Print BitShiftAriStrBin("00101011", 1, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '00010101
            Debug.Print BitShiftAriStrBin("00101011", 2, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '00001010
            Debug.Print BitShiftAriStrBin("00101011", 3, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '00000101
            Debug.Print BitShiftAriStrBin("00101011", 4, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '00000010
            Debug.Print BitShiftAriStrBin("00101011", 5, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '00000001
            Debug.Print BitShiftAriStrBin("00101011", 6, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '00000000
            Debug.Print BitShiftAriStrBin("00101011", 7, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '00000000
            Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '00101011
            Debug.Print BitShiftAriStrBin("00101011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '0000000001010110
            Debug.Print BitShiftAriStrBin("00101011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '0000000010101100
            Debug.Print BitShiftAriStrBin("00101011", 5, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '0000010101100000
            Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 16)  '0000000000101011
            Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '00101011
            Debug.Print BitShiftAriStrBin("00101011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '01010110
            Debug.Print BitShiftAriStrBin("00101011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '00101100
            Debug.Print BitShiftAriStrBin("00101011", 3, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '01011000
            Debug.Print BitShiftAriStrBin("00101011", 4, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '00110000
            Debug.Print BitShiftAriStrBin("00101011", 5, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '01100000
            Debug.Print BitShiftAriStrBin("00101011", 6, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '01000000
            Debug.Print BitShiftAriStrBin("00101011", 7, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '00000000
            Debug.Print "<<test 02>>"
            Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '10101011
            Debug.Print BitShiftAriStrBin("10101011", 1, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '11010101
            Debug.Print BitShiftAriStrBin("10101011", 2, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '11101010
            Debug.Print BitShiftAriStrBin("10101011", 3, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '11110101
            Debug.Print BitShiftAriStrBin("10101011", 4, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '11111010
            Debug.Print BitShiftAriStrBin("10101011", 5, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '11111101
            Debug.Print BitShiftAriStrBin("10101011", 6, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '11111110
            Debug.Print BitShiftAriStrBin("10101011", 7, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)  '11111111
            Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 16) '1111111110101011
            Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '10101011
            Debug.Print BitShiftAriStrBin("10101011", 1, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '11010101
            Debug.Print BitShiftAriStrBin("10101011", 2, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '11101010
            Debug.Print BitShiftAriStrBin("10101011", 3, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '11110101
            Debug.Print BitShiftAriStrBin("10101011", 4, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '11111010
            Debug.Print BitShiftAriStrBin("10101011", 5, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '11111101
            Debug.Print BitShiftAriStrBin("10101011", 6, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '11111110
            Debug.Print BitShiftAriStrBin("10101011", 7, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)  '11111111
            Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '10101011
            Debug.Print BitShiftAriStrBin("10101011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '1111111101010110
            Debug.Print BitShiftAriStrBin("10101011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '1111111010101100
            Debug.Print BitShiftAriStrBin("10101011", 5, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 0)   '1111010101100000
            Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 16)  '1111111110101011
            Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '10101011
            Debug.Print BitShiftAriStrBin("10101011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '11010110
            Debug.Print BitShiftAriStrBin("10101011", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '10101100
            Debug.Print BitShiftAriStrBin("10101011", 3, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '11011000
            Debug.Print BitShiftAriStrBin("10101011", 4, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '10110000
            Debug.Print BitShiftAriStrBin("10101011", 5, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '11100000
            Debug.Print BitShiftAriStrBin("10101011", 6, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '11000000
            Debug.Print BitShiftAriStrBin("10101011", 7, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '10000000
            Debug.Print BitShiftAriStrBin("10101011", 8, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '10000000
        '   Debug.Print "<<test 03>>"
        '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 8)   '�v���O������~
        '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 5)   '�v���O������~
        '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 4)   '�v���O������~
        '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 3)   '�v���O������~
        '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 2)   '�v���O������~
        '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 1)   '�v���O������~
        '   Debug.Print BitShiftAriStrBin("10001011", 1, 5, ARITHMETIC_SHIFT_SIGNBITSAVE, 10)           '�v���O������~
        '   Debug.Print BitShiftAriStrBin("", 2, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10)          '�v���O������~
        '   Debug.Print BitShiftAriStrBin("10001011", -1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10) '�v���O������~
        '   Debug.Print BitShiftAriStrBin("1000101A", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, 10)  '�v���O������~
        '   Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE, -1)  '�v���O������~
        End If
        Debug.Print "*** test finished! ***"
    End Sub

