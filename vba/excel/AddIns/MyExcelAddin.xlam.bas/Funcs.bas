Attribute VB_Name = "Funcs"
Option Explicit

' my excel addin functions v1.13

' ==================================================================
' =  <<�֐��ꗗ>>
' =    TextJoin2            �w�肵���͈͂̕��������������B
' =    TextSplit            ������𕪊����A�w�肵���v�f�̕������ԋp����B
' =    GetStrNum            �w�蕶����̌���ԋp����B
' =
' =    RemoveTailWord       ������؂蕶���ȍ~�̕��������������
' =    ExtractTailWord      ������؂蕶���ȍ~�̕������ԋp����
' =    GetDirPath           �w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����B
' =    GetFileName          �w�肳�ꂽ�t�@�C���p�X����t�@�C�����𒊏o����B
' =    GetFileExt           �w�肳�ꂽ�t�@�C���p�X����g���q�𒊏o����
' =    GetFileBase          �w�肳�ꂽ�t�@�C���p�X����t�@�C���x�[�X���𒊏o����
' =    GetFilePart          �w�肳�ꂽ�t�@�C���p�X����w�肳�ꂽ�ꕔ�𒊏o����
' =
' =    GetStrikeExist       ���������̗L���𔻒肷��B
' =    GetFontColor         �t�H���g�J���[��ԋp����B(�F�w���)
' =    GetInteriorColor     �w�i�F��ԋp����B(�F�w���)
' =    GetFontColorAll      �t�H���g�J���[��ԋp����B(�S�F�w���)
' =    GetInteriorColorAll  �w�i�F��ԋp����B(�S�F�w���)
' =    GetURL               �n�C�p�[�����N����URL�𒊏o����B
' =
' =    BitAndVal            �r�b�g�`�m�c���Z���s���B�i���l�j
' =    BitAndStrHex         �r�b�g�`�m�c���Z���s���B�i������P�U�i���j
' =    BitAndStrBin         �r�b�g�`�m�c���Z���s���B�i������Q�i���j
' =    BitOrVal             �r�b�g�n�q���Z���s���B�i���l�j
' =    BitOrStrHex          �r�b�g�n�q���Z���s���B�i������P�U�i���j
' =    BitOrStrBin          �r�b�g�n�q���Z���s���B�i������Q�i���j
' =    BitShiftVal          �r�b�g�r�g�h�e�s���Z���s���B�i���l�j
' =    BitShiftStrHex       �r�b�g�r�g�h�e�s���Z���s���B�i������P�U�i���j
' =    BitShiftStrBin       �r�b�g�r�g�h�e�s���Z���s���B�i������Q�i���j
' =
' =    RegExpSearch         ���K�\���������s���B
' =
' =    ConvSnakeToPascal    �����K���ϊ����s���i�X�l�[�N�P�[�X�˃p�X�J���P�[�X�j
' =    ConvSnakeToCamel     �����K���ϊ����s���i�X�l�[�N�P�[�X�˃L�������P�[�X�j
' =    ConvCamelToSnake     �����K���ϊ����s���i�L�������P�[�X�˃X�l�[�N�P�[�X�j
' =
' =    DiffRange            �w�肵���Q�͈̔͂��r���āA���S��v���ǂ����𔻒肷��
' =
' =    CalcPaddingWidth     �����񒷂����؂蕝�ʒu�܂ł̕�������ԋp����
' =    CalcPaddingTabWidth  �����񒷂���^�u����؂�ʒu�܂ł̕�������ԋp����
' =
' =    Exists               �t�@�C��/�t�H���_�̑��݊m�F���s��
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
' = ����    rConcRange      Range   [in]  ��������͈�
' = ����    sDlmtr          String  [in]  ��؂蕶���i�ȗ��j
' = ����    bIgnoreBlanc    Boolean [in]  �󔒖����i�ȗ��j
' = �ߒl                    Variant       ������̕�����
' = �o��    �E1�s��������1����w�肷�邱��
' = �ˑ�    �Ȃ�
' = ����    Funcs.bas
' ==================================================================
Public Function TextJoin2( _
    ByRef rConcRange As Range, _
    Optional ByVal sDlmtr As String = "", _
    Optional ByVal bIgnoreBlanc As Boolean = True _
) As Variant
    Dim rConcRangeCnt As Range
    Dim sConcTxtBuf As String
    
    If rConcRange Is Nothing Then
        TextJoin2 = CVErr(xlErrRef)  '�G���[�l
    Else
        If rConcRange.Rows.Count = 1 Or _
           rConcRange.Columns.Count = 1 Then
           
            ' ���s����
            If sDlmtr = "\n" Then
                sDlmtr = vbLf
            ElseIf sDlmtr = "\r" Then
                sDlmtr = vbCr
            ElseIf sDlmtr = "\r\n" Then
                sDlmtr = vbCrLf
            Else
                'Do Nothing
            End If
            
            If bIgnoreBlanc = True Then
                For Each rConcRangeCnt In rConcRange
                    If rConcRangeCnt.Value <> "" Then
                        sConcTxtBuf = sConcTxtBuf & sDlmtr & rConcRangeCnt.Value
                    End If
                Next rConcRangeCnt
            Else
                For Each rConcRangeCnt In rConcRange
                    sConcTxtBuf = sConcTxtBuf & sDlmtr & rConcRangeCnt.Value
                Next rConcRangeCnt
            End If
            
            ' ��؂蕶������
            If sDlmtr <> "" Then
                TextJoin2 = Mid$(sConcTxtBuf, Len(sDlmtr) + 1)
            Else
                TextJoin2 = sConcTxtBuf
            End If
        Else
            TextJoin2 = CVErr(xlErrRef)  '�G���[�l
        End If
    End If
End Function
    Private Sub Test_TextJoin2()
        Dim oTrgtRangePos01 As Range
        Dim oTrgtRangePos02 As Range
        Dim oTrgtRangePos03 As Range
        Dim oTrgtRangeNeg01 As Range
        Dim oTrgtRangeNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtRangePos01 = .Cells(1, 1)
            Set oTrgtRangePos02 = .Range(.Cells(1, 1), .Cells(1, 3))
            Set oTrgtRangePos03 = .Range(.Cells(1, 3), .Cells(1, 1))
            Set oTrgtRangeNeg01 = .Range(.Cells(1, 1), .Cells(3, 3))
            Set oTrgtRangeNeg02 = Nothing
        End With
        
        Dim lIdx
        Dim asStrBefore() As String
        For lIdx = 0 To oTrgtRangeNeg01.Count
            ReDim Preserve asStrBefore(lIdx)
            asStrBefore(lIdx) = oTrgtRangeNeg01.Item(lIdx + 1)
        Next lIdx
        
        Debug.Print "*** test start! ***"
        oTrgtRangePos01.Item(1) = "aaa"
        Debug.Print TextJoin2(oTrgtRangePos01, "\") 'aaa
        Debug.Print TextJoin2(oTrgtRangePos01, "")  'aaa
        oTrgtRangePos02.Item(1) = "bbb"
        oTrgtRangePos02.Item(2) = "ccc"
        oTrgtRangePos02.Item(3) = "ddd"
        Debug.Print TextJoin2(oTrgtRangePos02, "\")  'bbb\ccc\ddd
        Debug.Print TextJoin2(oTrgtRangePos02, "  ") 'bbb  ccc  ddd
        Debug.Print TextJoin2(oTrgtRangePos02, "")   'bbbcccddd
        Debug.Print TextJoin2(oTrgtRangePos02, "\n") 'bbb(���s)ccc(���s)ddd
        Debug.Print TextJoin2(oTrgtRangePos02, "\r") 'bbb(���s)ccc(���s)ddd
        Debug.Print TextJoin2(oTrgtRangePos02, "\r\n") 'bbb(���s)ccc(���s)ddd
        oTrgtRangePos03.Item(1) = "eee"
        oTrgtRangePos03.Item(2) = "fff"
        oTrgtRangePos03.Item(3) = "ggg"
        Debug.Print TextJoin2(oTrgtRangePos03, "\")  'eee\fff\ggg
        Debug.Print TextJoin2(oTrgtRangePos03, "  ") 'eee  fff  ggg
        Debug.Print TextJoin2(oTrgtRangePos03, "")   'eeefffggg
        Debug.Print TextJoin2(oTrgtRangeNeg01, "\")  '�G���[ 2023
        Debug.Print TextJoin2(oTrgtRangeNeg02, "\")  '�G���[ 2023
        Debug.Print "*** test finished! ***"
        
        For lIdx = 0 To oTrgtRangeNeg01.Count
            oTrgtRangeNeg01.Item(lIdx + 1) = asStrBefore(lIdx)
        Next lIdx
    End Sub

'Macros v1.10a �ȑO�Ƃ̌݊����ێ��p
Public Function ConcStr( _
    ByRef rConcRange As Range, _
    Optional ByVal sDlmtr As String = "", _
    Optional ByVal bIsBlancIgnore As Boolean = True _
) As Variant
    ConcStr = TextJoin2(rConcRange, sDlmtr, bIsBlancIgnore)
End Function

' ==================================================================
' = �T�v    ������𕪊����A�w�肵���v�f�̕������ԋp����
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = ����    iExtIndex   String  [in]  ���o����v�f ( 0 origin )
' = �ߒl                Variant       ���o������
' = �o��    iExtIndex ���v�f�𒴂���ꍇ�A�󕶎����ԋp����
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function TextSplit( _
    ByVal sStr As String, _
    ByVal sDlmtr As String, _
    ByVal iExtIndex As Integer _
) As Variant
    If sDlmtr = "" Then
        TextSplit = sStr
    Else
        If sStr = "" Then
            TextSplit = ""
        Else
            Dim vSplitStr As Variant
            vSplitStr = Split(sStr, sDlmtr) ' �����񕪊�
            If iExtIndex > UBound(vSplitStr) Or _
               iExtIndex < LBound(vSplitStr) Then
                TextSplit = ""
            Else
                TextSplit = vSplitStr(iExtIndex)
            End If
        End If
    End If
End Function
    Private Sub Test_TextSplit()
        Debug.Print "*** test start! ***"
        Debug.Print TextSplit("c:\test\a.txt", "\", 0)  'c:
        Debug.Print TextSplit("c:\test\a.txt", "\", 1)  'test
        Debug.Print TextSplit("c:\test\a.txt", "\", 2)  'a.txt
        Debug.Print TextSplit("c:\test\a.txt", "\", -1) '
        Debug.Print TextSplit("c:\test\a.txt", "\", 3)  '
        Debug.Print TextSplit("", "\", 1)               '
        Debug.Print TextSplit("c:\a.txt", "", 1)        'c:\a.txt
        Debug.Print TextSplit("", "", 1)                '
        Debug.Print TextSplit("", "", 0)                '
        Debug.Print "*** test finished! ***"
    End Sub

'Macros v1.10a �ȑO�Ƃ̌݊����ێ��p
Public Function SplitStr( _
    ByVal sStr As String, _
    ByVal sDlmtr As String, _
    ByVal iExtIndex As Integer _
) As Variant
    SplitStr = TextSplit(sStr, sDlmtr, iExtIndex)
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
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
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
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
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
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & ExtractTailWord("", "\")               '
        Result = Result & vbNewLine & ExtractTailWord("c:\a", "\")           ' a
        Result = Result & vbNewLine & ExtractTailWord("c:\a\", "\")          '
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b", "\")         ' b
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\", "\")        '
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "\")   ' c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\\b\c.txt", "\")    ' c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
        Result = Result & vbNewLine & ExtractTailWord("c:\a\\b\c.txt", "\\") ' b\c.txt
        Result = Result & vbNewLine & "*** test finished! ***"
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕��������������B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ����������
' = �o��    �Ȃ�
' = �ˑ�    Mng_String.bas/ExtractTailWord()
' = ����    Mng_String.bas
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
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & RemoveTailWord("", "\")               '
        Result = Result & vbNewLine & RemoveTailWord("c:\a", "\")           ' c:
        Result = Result & vbNewLine & RemoveTailWord("c:\a\", "\")          ' c:\a
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b", "\")         ' c:\a
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\", "\")        ' c:\a\b
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "\")   ' c:\a\b
        Result = Result & vbNewLine & RemoveTailWord("c:\\b\c.txt", "\")    ' c:\\b
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "")    ' c:\a\b\c.txt
        Result = Result & vbNewLine & RemoveTailWord("c:\a\b\c.txt", "\\")  ' c:\a\b\c.txt
        Result = Result & vbNewLine & RemoveTailWord("c:\a\\b\c.txt", "\\") ' c:\a
        Result = Result & vbNewLine & "*** test finished! ***"
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����
' = ����    sFilePath       String  [in]  �t�@�C���p�X
' = ����    bErrorEnable    Boolean [in]  �G���[�����L��/����(��)
' = �ߒl                    Variant       �t�H���_�p�X
' = �o��    ���[�J���t�@�C���p�X�i��Fc:\test�j�� URL �i��Fhttps://test�j
' =         ���w��\�B
' =         (��) bErrorEnable �ɂăt�@�C���p�X�ȊO���w�肳�ꂽ���̕ԋp�l��
' =         �ς��邱�Ƃ��o����
' =            True  : sFilePath ��ԋp
' =            False : �G���[�l�ixlErrNA�j��ԋp
' = �ˑ�    Mng_String.bas/RemoveTailWord()
' = ����    Mng_String.bas
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath As String, _
    Optional ByVal bErrorEnable As Boolean = False _
) As Variant
    If InStr(sFilePath, "\") Then
        GetDirPath = RemoveTailWord(sFilePath, "\")
    ElseIf InStr(sFilePath, "/") Then
        GetDirPath = RemoveTailWord(sFilePath, "/")
    Else
        If bErrorEnable = True Then
            GetDirPath = CVErr(xlErrNA)  '�G���[�l
        Else
            GetDirPath = sFilePath
        End If
    End If
End Function
    Private Sub Test_GetDirPath()
        Dim Result As String
        Dim vRet As Variant
        Result = "[Result]"
        vRet = GetDirPath("C:\test\a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' C:\test
        vRet = GetDirPath("http://test/a", True): Result = Result & vbNewLine & CStr(vRet)  ' http://test
        vRet = GetDirPath("C:_test_a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' �G���[ 2042
        Result = Result & vbNewLine                                                         '
        vRet = GetDirPath("C:\test\a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' C:\test
        vRet = GetDirPath("http://test/a", False): Result = Result & vbNewLine & CStr(vRet) ' http://test
        vRet = GetDirPath("C:_test_a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' C:_test_a.txt
        Result = Result & vbNewLine                                                         '
        vRet = GetDirPath("C:\test\a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' C:\test
        vRet = GetDirPath("http://test/a"): Result = Result & vbNewLine & CStr(vRet)        ' http://test
        vRet = GetDirPath("C:_test_a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' C:_test_a.txt
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�@�C�����𒊏o����
' = ����    sFilePath       String  [in]  �t�@�C���p�X
' = ����    bErrorEnable    Boolean [in]  �G���[�����L��/����(��)
' = �ߒl                    Variant       �t�@�C����
' = �o��    ���[�J���t�@�C���p�X�i��Fc:\test�j�� URL �i��Fhttps://test�j
' =         ���w��\�B
' =         (��) bErrorEnable �ɂăt�@�C���p�X�ȊO���w�肳�ꂽ���̕ԋp�l��
' =         �ς��邱�Ƃ��o����
' =            True  : sFilePath ��ԋp
' =            False : �G���[�l�ixlErrNA�j��ԋp
' = �ˑ�    Mng_String.bas/ExtractTailWord()
' = ����    Mng_String.bas
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath As String, _
    Optional ByVal bErrorEnable As Boolean = False _
) As Variant
    If InStr(sFilePath, "\") Then
        GetFileName = ExtractTailWord(sFilePath, "\")
    ElseIf InStr(sFilePath, "/") Then
        GetFileName = ExtractTailWord(sFilePath, "/")
    Else
        If bErrorEnable = True Then
            GetFileName = CVErr(xlErrNA)  '�G���[�l
        Else
            GetFileName = sFilePath
        End If
    End If
End Function
    Private Sub Test_GetFileName()
        Dim Result As String
        Dim vRet As Variant
        Result = "[Result]"
        vRet = GetFileName("C:\test\a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' a.txt
        vRet = GetFileName("http://test/a", True): Result = Result & vbNewLine & CStr(vRet)  ' a
        vRet = GetFileName("C:_test_a.txt", True): Result = Result & vbNewLine & CStr(vRet)  ' �G���[ 2042
        Result = Result & vbNewLine                                                          '
        vRet = GetFileName("C:\test\a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' a.txt
        vRet = GetFileName("http://test/a", False): Result = Result & vbNewLine & CStr(vRet) ' a
        vRet = GetFileName("C:_test_a.txt", False): Result = Result & vbNewLine & CStr(vRet) ' c:_test_a
        Result = Result & vbNewLine                                                          '
        vRet = GetFileName("C:\test\a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' a.txt
        vRet = GetFileName("http://test/a"): Result = Result & vbNewLine & CStr(vRet)        ' a
        vRet = GetFileName("C:_test_a.txt"): Result = Result & vbNewLine & CStr(vRet)        ' c:_test_a
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����g���q�𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �g���q
' = �o��    �E�g���q���Ȃ��ꍇ�A�󕶎���ԋp����
' =         �E�t�@�C�������w��\
' = �ˑ�    Mng_String.bas/GetFileName()
' =         Mng_String.bas/ExtractTailWord()
' = ����    Mng_String.bas
' ==================================================================
Public Function GetFileExt( _
    ByVal sFilePath As String _
) As String
    Dim sFileName As String
    sFileName = GetFileName(sFilePath)
    If InStr(sFileName, ".") > 0 Then
        GetFileExt = ExtractTailWord(sFileName, ".")
    Else
        GetFileExt = ""
    End If
End Function
    Private Sub Test_GetFileExt()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileExt("c:\codes\test.txt")     'txt
        Result = Result & vbNewLine & GetFileExt("c:\codes\test")         '
        Result = Result & vbNewLine & GetFileExt("test.txt")              'txt
        Result = Result & vbNewLine & GetFileExt("test")                  '
        Result = Result & vbNewLine & GetFileExt("c:\codes\test.aaa.txt") 'txt
        Result = Result & vbNewLine & GetFileExt("test.aaa.txt")          'txt
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�@�C���x�[�X���𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�@�C���x�[�X��
' = �o��    �E�g���q���Ȃ��ꍇ�A�󕶎���ԋp����
' =         �E�t�@�C�������w��\
' = �ˑ�    Mng_String.bas/GetFileName()
' =         Mng_String.bas/RemoveTailWord()
' = ����    Mng_String.bas
' ==================================================================
Public Function GetFileBase( _
    ByVal sFilePath As String _
) As String
    Dim sFileName As String
    sFileName = GetFileName(sFilePath)
    GetFileBase = RemoveTailWord(sFileName, ".")
End Function
    Private Sub Test_GetFileBase()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileBase("c:\codes\test.txt")     'test
        Result = Result & vbNewLine & GetFileBase("c:\codes\test")         'test
        Result = Result & vbNewLine & GetFileBase("test.txt")              'test
        Result = Result & vbNewLine & GetFileBase("test")                  'test
        Result = Result & vbNewLine & GetFileBase("c:\codes\test.aaa.txt") 'test.aaa
        Result = Result & vbNewLine & GetFileBase("test.aaa.txt")          'test.aaa
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����w�肳�ꂽ�ꕔ�𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = ����    lPartType   Long    [in]  ���o���
' =                                     1) �t�H���_�p�X
' =                                     2) �t�@�C����
' =                                     3) �t�@�C���x�[�X��
' =                                     4) �t�@�C���g���q
' = �ߒl                String        ���o�����ꕔ
' = �o��    �E���o��ʂ�����Ă���ꍇ�A�󕶎���ԋp����
' = �ˑ�    Mng_String.bas/GetDirPath()
' =         Mng_String.bas/GetFileName()
' =         Mng_String.bas/GetFileBase)
' =         Mng_String.bas/GetFileExt()
' = ����    Mng_String.bas
' ==================================================================
Public Function GetFilePart( _
    ByVal sFilePath As String, _
    ByVal lPartType As Long _
) As String
    Select Case lPartType
        Case 1: GetFilePart = GetDirPath(sFilePath)
        Case 2: GetFilePart = GetFileName(sFilePath)
        Case 3: GetFilePart = GetFileBase(sFilePath)
        Case 4: GetFilePart = GetFileExt(sFilePath)
        Case Else: GetFilePart = ""
    End Select
End Function
    Private Sub Test_GetFilePart()
        Dim Result As String
        Result = "[Result]"
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 0)     '
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 1)     'c:\codes
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 2)     'test.txt
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 3)     'test
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 4)     'txt
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 5)     '
        Debug.Print Result
    End Sub

' ==================================================================
' = �T�v    ���������̗L���𔻒肷�� (TRUE:�L�AFALSE:��)
' = ����    rRange   Range     [in]  �Z��
' = �ߒl             Variant         ���������L��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Funcs.bas
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
        Dim oTrgtRangePos01 As Range
        Dim oTrgtRangeNeg01 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtRangePos01 = .Cells(1, 1)
            Set oTrgtRangeNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
        End With
        
        Dim lStrkThrghBefore As Long
        lStrkThrghBefore = oTrgtRangePos01.Font.Strikethrough
        
        Debug.Print "*** test start! ***"
        oTrgtRangePos01.Font.Strikethrough = True
        Debug.Print GetStrikeExist(oTrgtRangePos01) 'True
        oTrgtRangePos01.Font.Strikethrough = False
        Debug.Print GetStrikeExist(oTrgtRangePos01) 'False
        Debug.Print GetStrikeExist(oTrgtRangeNeg01) '�G���[ 2023
        Debug.Print "*** test finished! ***"
        
        oTrgtRangePos01.Font.Strikethrough = lStrkThrghBefore
    End Sub

' ==================================================================
' = �T�v    �t�H���g�J���[��ԋp����(�F�w���)
' = ����    rTrgtRange  Range     [in]  �Z��
' = ����    sColorType  String    [in]  �F��ʁiR or G or B�j
' = ����    bIsHex      Boolean   [in]  ��i0:Decimal�A1:Hex�j(�ȗ���)
' = �ߒl                Variant         �t�H���g�F
' = �o��    �Ȃ�
' = �ˑ�    Funcs.bas/GetColor()
' = ����    Funcs.bas
' ==================================================================
Public Function GetFontColor( _
    ByRef rTrgtCell As Range, _
    ByVal sColorType As String, _
    Optional ByVal bIsHex As Boolean _
) As Variant
    GetFontColor = GetColor(rTrgtCell, sColorType, 1, bIsHex)
End Function
    Private Sub Test_GetFontColor()
        Dim oTrgtCellsPos01 As Range
        Dim oTrgtCellsNeg01 As Range
        Dim oTrgtCellsNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtCellsPos01 = .Cells(1, 1)
            Set oTrgtCellsNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
            Set oTrgtCellsNeg02 = .Range(.Cells(1, 1), .Cells(4, 1))
        End With
        
        Dim lColorBefore As Long
        lColorBefore = oTrgtCellsPos01.Font.Color
        
        Debug.Print "*** test start! ***"
        oTrgtCellsPos01.Font.Color = RGB(0, 0, 0)
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", False)  '0
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", False)  '0
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", False)  '0
        oTrgtCellsPos01.Font.Color = RGB(100, 100, 100)
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", False)  '100
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", False)  '100
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", False)  '100
        oTrgtCellsPos01.Font.Color = RGB(255, 255, 255)
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", False)  '255
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", False)  '255
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", False)  '255
        oTrgtCellsPos01.Font.Color = RGB(16, 100, 152)
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", False)  '16
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", False)  '100
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", False)  '152
        Debug.Print GetFontColor(oTrgtCellsPos01, "r", True)   '10
        Debug.Print GetFontColor(oTrgtCellsPos01, "g", True)   '64
        Debug.Print GetFontColor(oTrgtCellsPos01, "b", True)   '98
        Debug.Print GetFontColor(oTrgtCellsPos01, "", False)   '�G���[ 2015
        Debug.Print GetFontColor(oTrgtCellsPos01, "aa", False) '�G���[ 2015
        Debug.Print GetFontColor(oTrgtCellsNeg01, "r", False)  '�G���[ 2023
        Debug.Print GetFontColor(oTrgtCellsNeg02, "r", False)  '�G���[ 2023
        Debug.Print GetFontColor(oTrgtCellsPos01, "r")         '16
        Debug.Print "*** test finished! ***"
        
        oTrgtCellsPos01.Font.Color = lColorBefore
    End Sub

' ==================================================================
' = �T�v    �t�H���g�J���[��ԋp����(�S�F�w���)
' = ����    rTrgtRange  Range     [in]  �Z��
' = ����    bIsHex      Boolean   [in]  ��i0:Decimal�A1:Hex�j(�ȗ���)
' = �ߒl                Variant         �t�H���g�F
' = �o��    �Ȃ�
' = �ˑ�    Funcs.bas/GetColor()
' = ����    Funcs.bas
' ==================================================================
Public Function GetFontColorAll( _
    ByRef rTrgtCell As Range, _
    Optional ByVal bIsHex As Boolean = False _
) As Variant
    GetFontColorAll = GetColorAll(rTrgtCell, 1, bIsHex)
End Function
    Private Sub Test_GetFontColorAll()
        Dim oTrgtCellsPos01 As Range
        Dim oTrgtCellsNeg01 As Range
        Dim oTrgtCellsNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtCellsPos01 = .Cells(1, 1)
            Set oTrgtCellsNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
            Set oTrgtCellsNeg02 = .Range(.Cells(1, 1), .Cells(4, 1))
        End With
        
        Dim lColorBefore As Long
        lColorBefore = oTrgtCellsPos01.Font.Color
        
        Debug.Print "*** test start! ***"
        oTrgtCellsPos01.Font.Color = RGB(0, 0, 0)
        Debug.Print GetFontColorAll(oTrgtCellsPos01, False)  '0,0,0
        oTrgtCellsPos01.Font.Color = RGB(100, 100, 100)
        Debug.Print GetFontColorAll(oTrgtCellsPos01, False)  '100,100,100
        oTrgtCellsPos01.Font.Color = RGB(255, 255, 255)
        Debug.Print GetFontColorAll(oTrgtCellsPos01, False)  '255,255,255
        oTrgtCellsPos01.Font.Color = RGB(16, 100, 152)
        Debug.Print GetFontColorAll(oTrgtCellsPos01, False)  '16,100,152
        Debug.Print GetFontColorAll(oTrgtCellsPos01, True)   '10,64,98
        Debug.Print GetFontColorAll(oTrgtCellsNeg01, False)  '�G���[ 2023
        Debug.Print GetFontColorAll(oTrgtCellsNeg02, False)  '�G���[ 2023
        Debug.Print GetFontColorAll(oTrgtCellsPos01)         '16,100,152
        Debug.Print "*** test finished! ***"
        
        oTrgtCellsPos01.Font.Color = lColorBefore
    End Sub

' ==================================================================
' = �T�v    �w�i�F��ԋp����(�F�w���)
' = ����    rTrgtRange  Range     [in]  �Z��
' = ����    sColorType  String    [in]  �F��ʁiR or G or B�j
' = ����    bIsHex      Boolean   [in]  ��i0:Decimal�A1:Hex�j
' = �ߒl                Variant         �w�i�F
' = �o��    �Ȃ�
' = �ˑ�    Funcs.bas/ConvRgb2X()
' = ����    Funcs.bas
' ==================================================================
Public Function GetInteriorColor( _
    ByRef rTrgtCell As Range, _
    ByVal sColorType As String, _
    ByVal bIsHex As Boolean _
) As Variant
    GetInteriorColor = GetColor(rTrgtCell, sColorType, 2, bIsHex)
End Function
    Private Sub Test_GetInteriorColor()
        Dim oTrgtCellsPos01 As Range
        Dim oTrgtCellsNeg01 As Range
        Dim oTrgtCellsNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtCellsPos01 = .Cells(1, 1)
            Set oTrgtCellsNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
            Set oTrgtCellsNeg02 = .Range(.Cells(1, 1), .Cells(4, 1))
        End With
        
        Dim lColorBefore As Long
        lColorBefore = oTrgtCellsPos01.Interior.Color
        
        Debug.Print "*** test start! ***"
        oTrgtCellsPos01.Interior.Color = RGB(0, 0, 0)
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", False)  '0
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", False)  '0
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", False)  '0
        oTrgtCellsPos01.Interior.Color = RGB(100, 100, 100)
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", False)  '100
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", False)  '100
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", False)  '100
        oTrgtCellsPos01.Interior.Color = RGB(255, 255, 255)
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", False)  '255
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", False)  '255
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", False)  '255
        oTrgtCellsPos01.Interior.Color = RGB(16, 100, 152)
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", False)  '16
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", False)  '100
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", False)  '152
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "r", True)   '10
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "g", True)   '64
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "b", True)   '98
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "", False)   '�G���[ 2015
        Debug.Print GetInteriorColor(oTrgtCellsPos01, "aa", False) '�G���[ 2015
        Debug.Print GetInteriorColor(oTrgtCellsNeg01, "r", False)  '�G���[ 2023
        Debug.Print GetInteriorColor(oTrgtCellsNeg02, "r", False)  '�G���[ 2023
        Debug.Print "*** test finished! ***"
        
        oTrgtCellsPos01.Interior.Color = lColorBefore
    End Sub

' ==================================================================
' = �T�v    �w�i�F��ԋp����(�S�F�w���)
' = ����    rTrgtRange  Range     [in]  �Z��
' = ����    bIsHex      Boolean   [in]  ��i0:Decimal�A1:Hex�j
' = �ߒl                Variant         �w�i�F
' = �o��    �Ȃ�
' = �ˑ�    Funcs.bas/GetColorAll()
' = ����    Funcs.bas
' ==================================================================
Public Function GetInteriorColorAll( _
    ByRef rTrgtCell As Range, _
    Optional ByVal bIsHex As Boolean = False _
) As Variant
    GetInteriorColorAll = GetColorAll(rTrgtCell, 2, bIsHex)
End Function
    Private Sub Test_GetInteriorColorAll()
        Dim oTrgtCellsPos01 As Range
        Dim oTrgtCellsNeg01 As Range
        Dim oTrgtCellsNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtCellsPos01 = .Cells(1, 1)
            Set oTrgtCellsNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
            Set oTrgtCellsNeg02 = .Range(.Cells(1, 1), .Cells(4, 1))
        End With
        
        Dim lColorBefore As Long
        lColorBefore = oTrgtCellsPos01.Interior.Color
        
        Debug.Print "*** test start! ***"
        oTrgtCellsPos01.Interior.Color = RGB(0, 0, 0)
        Debug.Print GetInteriorColorAll(oTrgtCellsPos01, False)  '0,0,0
        oTrgtCellsPos01.Interior.Color = RGB(100, 100, 100)
        Debug.Print GetInteriorColorAll(oTrgtCellsPos01, False)  '100,100,100
        oTrgtCellsPos01.Interior.Color = RGB(255, 255, 255)
        Debug.Print GetInteriorColorAll(oTrgtCellsPos01, False)  '255,255,255
        oTrgtCellsPos01.Interior.Color = RGB(16, 100, 152)
        Debug.Print GetInteriorColorAll(oTrgtCellsPos01, False)  '16,100,152
        Debug.Print GetInteriorColorAll(oTrgtCellsPos01, True)   '10,64,98
        Debug.Print GetInteriorColorAll(oTrgtCellsNeg01, False)  '�G���[ 2023
        Debug.Print GetInteriorColorAll(oTrgtCellsNeg02, False)  '�G���[ 2023
        Debug.Print GetInteriorColorAll(oTrgtCellsPos01)         '16,100,152
        Debug.Print "*** test finished! ***"
        
        oTrgtCellsPos01.Interior.Color = lColorBefore
    End Sub

' ==================================================================
' = �T�v    �n�C�p�[�����N����URL�𒊏o����B
' = ����    rTrgtRange  Range     [in]  �Z��
' = �ߒl                String          URL
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Funcs.bas
' ==================================================================
Public Function GetURL( _
    ByRef rTrgtCell As Range _
) As String
    Dim sAddr As String
    If rTrgtCell.Hyperlinks.Count > 0 Then
        If rTrgtCell.Hyperlinks(1).Address Like "http*" Then
            sAddr = rTrgtCell.Hyperlinks(1).Address
        End If
    End If
    If sAddr <> "" Then
        GetURL = sAddr
    Else
        GetURL = ""
    End If
End Function
    Private Sub Test_GetURL()
        'TODO
    End Sub

' ==================================================================
' = �T�v    �r�b�g�`�m�c���Z���s���B�i���l�j
' = ����    cInVal1   Currency   [in]  ���͒l �����i10�i�����l�j
' = ����    cInVal2   Currency   [in]  ���͒l �E���i10�i�����l�j
' = �ߒl              Variant          ���Z���ʁi10�i�����l�j
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function BitAndVal( _
    ByVal cInVal1 As Currency, _
    ByVal cInVal2 As Currency _
) As Variant
    If cInVal1 > 2147483647# Or cInVal1 < -2147483647# Or _
       cInVal2 > 2147483647# Or cInVal2 < -2147483647# Then
        BitAndVal = CVErr(xlErrNum)  '�G���[�l
    Else
        BitAndVal = cInVal1 And cInVal2
    End If
End Function
    Private Sub Test_BitAndVal()
        Debug.Print "*** test start! ***"
        Debug.Print BitAndVal(&HFFFF&, &HFF00&)         '65280 (0xFF00)
        Debug.Print BitAndVal(&HFFFF&, &HFF&)           '255 (0xFF)
        Debug.Print BitAndVal(&HFFFF&, &HA5A5&)         '42405 (0xA5A5)
        Debug.Print BitAndVal(&HA5&, &HA500&)           '0
        Debug.Print BitAndVal(&H1&, &H8&)               '0
        Debug.Print BitAndVal(&H1&, &HA&)               '0
        Debug.Print BitAndVal(&H5&, &HA&)               '0
        Debug.Print BitAndVal(&H7FFFFFFF, &HFF&)        '255 (0xFF)
        Debug.Print BitAndVal(&H80000000, &HFF&)        '�G���[ 2036
        Debug.Print BitAndVal(2147483648#, &HFF&)       '�G���[ 2036
        Debug.Print BitAndVal(2147483647#, &HFF&)       '255 (0xFF)
        Debug.Print BitAndVal(-2147483647#, &HFF&)      '1
        Debug.Print BitAndVal(-2147483648#, &HFF&)      '�G���[ 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�`�m�c���Z���s���B�i������P�U�i���j
' = ����    sInHexVal1  String     [in]  ���͒l �����i������j
' = ����    sInHexVal2  String     [in]  ���͒l �E���i������j
' = ����    lInDigitNum Long       [in]  �o�͌���
' = �ߒl                Variant          ���Z���ʁi������j
' = �o��    �Ȃ�
' = �ˑ�    Mng_String.bas/Hex2Bin()
' =         Mng_String.bas/BitAndStrBin()
' =         Mng_String.bas/Bin2Hex()
' = ����    Mng_String.bas
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
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
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
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function BitOrVal( _
    ByVal cInVal1 As Currency, _
    ByVal cInVal2 As Currency _
) As Variant
    Dim sHexVal As String
    If cInVal1 > 2147483647# Or cInVal1 < -2147483647# Or _
       cInVal2 > 2147483647# Or cInVal2 < -2147483647# Then
        BitOrVal = CVErr(xlErrNum)  '�G���[�l
    Else
        BitOrVal = cInVal1 Or cInVal2
    End If
End Function
    Private Sub Test_BitOrVal()
        Debug.Print "*** test start! ***"
        Debug.Print BitOrVal(&HFFFF&, &HFF00&)      '65535 (0xFFFF)
        Debug.Print BitOrVal(&HFFFF&, &HFF&)        '65535 (0xFFFF)
        Debug.Print BitOrVal(&HFFFF&, &HA5A5&)      '65535 (0xFFFF)
        Debug.Print BitOrVal(&HA5&, &HA500&)        '42405 (0xA5A5)
        Debug.Print BitOrVal(&H1&, &H8&)            '9
        Debug.Print BitOrVal(&H1&, &HA&)            '11 (0xB)
        Debug.Print BitOrVal(&H5&, &HA&)            '15 (0xF)
        Debug.Print BitOrVal(&H7FFFFFFF, &HFF&)     '2147483647 (0x7FFFFFFF)
        Debug.Print BitOrVal(&H80000000, &HFF&)     '�G���[ 2036
        Debug.Print BitOrVal(2147483648#, &HFF&)    '�G���[ 2036
        Debug.Print BitOrVal(2147483647#, &HFF&)    '2147483647 (0x7FFFFFFF)
        Debug.Print BitOrVal(-2147483647#, &HFF&)   '-2147483393 (0x800000FF)
        Debug.Print BitOrVal(-2147483648#, &HFF&)   '�G���[ 2036
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �r�b�g�n�q���Z���s���B�i������P�U�i���j
' = ����    sInHexVal1  String     [in]  ���͒l �����i������j
' = ����    sInHexVal2  String     [in]  ���͒l �E���i������j
' = ����    lInDigitNum Long       [in]  �o�͌���
' = �ߒl                Variant          ���Z���ʁi������j
' = �o��    �Ȃ�
' = �ˑ�    Mng_String.bas/Hex2Bin()
' =         Mng_String.bas/BitOrStrBin()
' =         Mng_String.bas/Bin2Hex()
' = ����    Mng_String.bas
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
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
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
' = �o��    32�r�b�g�̂ݑΉ�����B���̂��߁A���V�t�g�̌��ʂ�32�r�b�g��
' =         ������ꍇ�A����32�r�b�g�̃V�t�g���ʂ�ԋp����B
' = �ˑ�    Mng_String.bas/Dec2Hex()
' =         Mng_String.bas/BitShiftStrBin()
' =         Mng_String.bas/Hex2Bin()
' =         Mng_String.bas/Bin2Hex()
' =         Mng_String.bas/Hex2Dec()
' = ����    Mng_String.bas
' ==================================================================
Public Function BitShiftVal( _
    ByVal cInDecVal As Currency, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    Optional ByVal eInShiftType As E_SHIFT_TYPE = LOGICAL_SHIFT _
) As Variant
    If cInDecVal < -2147483648# Or cInDecVal > 4294967295# Then
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
    
    'Dec��Hex
    Dim sPreHexVal As String
    sPreHexVal = Dec2Hex(cInDecVal)
    If sPreHexVal = "error" Then
        BitShiftVal = CVErr(xlErrNum) '�G���[�l
        Exit Function
    End If
    
    'Hex��Bin
    Dim sPreBinVal As String
    sPreBinVal = Hex2Bin(sPreHexVal)
    Debug.Assert sPreBinVal <> "error"
    
    'Shift
    Dim sPostBinVal As String
    sPostBinVal = BitShiftStrBin(sPreBinVal, lInShiftNum, eInDirection, eInShiftType, 32)
    Debug.Assert sPostBinVal <> "error"
    
    'Bin��Hex
    Dim sPostHexVal As String
    sPostHexVal = Bin2Hex(sPostBinVal, True)
    Debug.Assert sPostHexVal <> "error"
    
    'Hex��Dec
    Dim vOutDecVal As Variant
    vOutDecVal = Hex2Dec(sPostHexVal, False)
    If vOutDecVal = "error" Then
        BitShiftVal = CVErr(xlErrNum) '�G���[�l
        Exit Function
    End If
    
    BitShiftVal = vOutDecVal
End Function
    Private Sub Test_BitShiftVal()
        Debug.Print "*** test start! ***"
        Debug.Print Hex(BitShiftVal(&H10&, 0, RIGHT_SHIFT, LOGICAL_SHIFT))          '10
        Debug.Print Hex(BitShiftVal(&H10&, 1, RIGHT_SHIFT, LOGICAL_SHIFT))          '8
        Debug.Print Hex(BitShiftVal(&H10&, 2, RIGHT_SHIFT, LOGICAL_SHIFT))          '4
        Debug.Print Hex(BitShiftVal(&H10&, 3, RIGHT_SHIFT, LOGICAL_SHIFT))          '2
        Debug.Print Hex(BitShiftVal(&H10&, 4, RIGHT_SHIFT, LOGICAL_SHIFT))          '1
        Debug.Print Hex(BitShiftVal(&H10&, 5, RIGHT_SHIFT, LOGICAL_SHIFT))          '0
        Debug.Print Hex(BitShiftVal(&H10&, 8, RIGHT_SHIFT, LOGICAL_SHIFT))          '0
        Debug.Print Hex(BitShiftVal(&H10&, 0, LEFT_SHIFT, LOGICAL_SHIFT))           '10
        Debug.Print Hex(BitShiftVal(&H10&, 1, LEFT_SHIFT, LOGICAL_SHIFT))           '20
        Debug.Print Hex(BitShiftVal(&H10&, 2, LEFT_SHIFT, LOGICAL_SHIFT))           '40
        Debug.Print Hex(BitShiftVal(&H10&, 3, LEFT_SHIFT, LOGICAL_SHIFT))           '80
        Debug.Print Hex(BitShiftVal(&H10&, 8, LEFT_SHIFT, LOGICAL_SHIFT))           '1000
        Debug.Print Hex(BitShiftVal(&H10&, 12, LEFT_SHIFT, LOGICAL_SHIFT))          '10000
        Debug.Print Hex(BitShiftVal(&H10&, 16, LEFT_SHIFT, LOGICAL_SHIFT))          '100000
        Debug.Print Hex(BitShiftVal(&H10&, 20, LEFT_SHIFT, LOGICAL_SHIFT))          '1000000
        Debug.Print Hex(BitShiftVal(&H10&, 24, LEFT_SHIFT, LOGICAL_SHIFT))          '10000000
        Debug.Print Hex(BitShiftVal(&H10&, 25, LEFT_SHIFT, LOGICAL_SHIFT))          '20000000
        Debug.Print Hex(BitShiftVal(&H10&, 26, LEFT_SHIFT, LOGICAL_SHIFT))          '40000000
       'Debug.Print Hex(BitShiftVal(&H10&, 27, LEFT_SHIFT, LOGICAL_SHIFT))          '�G���[�iHex()�ɂăI�[�o�[�t���[�j
        Debug.Print BitShiftVal(&H10&, 27, LEFT_SHIFT, LOGICAL_SHIFT)               '2147483648
        Debug.Print BitShiftVal(&H10&, 28, LEFT_SHIFT, LOGICAL_SHIFT)               '0
        Debug.Print BitShiftVal(&H10&, 29, LEFT_SHIFT, LOGICAL_SHIFT)               '0
        Debug.Print Hex(BitShiftVal(&H7FFFFFFF, 0, LEFT_SHIFT, LOGICAL_SHIFT))      '7FFFFFFF
        Debug.Print BitShiftVal(&H80000000, 0, LEFT_SHIFT, LOGICAL_SHIFT)           '2147483648 (0x80000000)
        Debug.Print BitShiftVal(4294967294#, 0, LEFT_SHIFT, LOGICAL_SHIFT)          '4294967294 (0xFFFFFFFE)
        Debug.Print BitShiftVal(&HFFFFFFFE, 0, LEFT_SHIFT, LOGICAL_SHIFT)           '4294967294 (0xFFFFFFFE)
        Debug.Print BitShiftVal(&H10&, 1, LEFT_SHIFT, ARITHMETIC_SHIFT_SIGNBITSAVE) '32
        Debug.Print BitShiftVal(&H10&, -1, LEFT_SHIFT, LOGICAL_SHIFT)               '�G���[ 2036
        Debug.Print BitShiftVal(&H10&, 1, 3, LOGICAL_SHIFT)                         '�G���[ 2015
        Debug.Print BitShiftVal(&H10&, 1, LEFT_SHIFT, 3)                            '�G���[ 2015
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
' = �ˑ�    Mng_String.bas/BitShiftStrBin()
' =         Mng_String.bas/Hex2Bin()
' =         Mng_String.bas/Bin2Hex()
' = ����    Mng_String.bas
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
' = �ˑ�    Mng_String.bas/BitShiftLogStrBin()
' =         Mng_String.bas/BitShiftAriStrBin()
' = ����    Mng_String.bas
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
            BitShiftStrBin = BitShiftLogStrBin(sInBinVal, lInShiftNum, eInDirection, lInBitLen)
        Case ARITHMETIC_SHIFT_SIGNBITSAVE:
            BitShiftStrBin = BitShiftAriStrBin(sInBinVal, lInShiftNum, eInDirection, lInBitLen, True, False)
        Case ARITHMETIC_SHIFT_SIGNBITTRUNC:
            BitShiftStrBin = BitShiftAriStrBin(sInBinVal, lInShiftNum, eInDirection, lInBitLen, False, False)
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
' = �T�v    ���K�\���������s���iExcel�֐��p�j
' = ����    sSearchPattern  String   [in]  �����p�^�[��
' = ����    sTargetStr      String   [in]  �����Ώە�����
' = ����    lMatchIdx       Long     [in]  �������ʃC���f�b�N�X�i�����ȗ��j
' = ����    bIsIgnoreCase   Boolean  [in]  ��/��������ʂ��Ȃ����i�����ȗ��j
' = ����    bIsGlobal       Boolean  [in]  ������S�̂��������邩�i�����ȗ��j
' = �ߒl                    Variant        ��������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
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
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
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
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
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
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
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

' ==================================================================
' = �T�v    �w�肵���Q�͈̔͂��r���āA���S��v���ǂ����𔻒肷��
' = ����    rTrgtRange01        Range   [in]  ��r�Ώ۔͈͂P
' = ����    rTrgtRange02        Range   [in]  ��r�Ώ۔͈͂Q
' = ����    bCellPosCheckValid  Boolean [in]  �Z���ʒu�`�F�b�N�L��/����
' = �ߒl                        Boolean       ��r����
' = �o��    �ȉ��̂����ꂩ�𖞂����ꍇ�AFalse ��ԋp����
' =           �E�͈͓��̃Z�������s��v
' =           �E�͈͓��̍s�����s��v
' =           �E�͈͓��̗񐔂��s��v
' =           �E�͈͓��̊e�Z���̒l���s��v
' =           �E�͈͓��̊J�n�Z���Ɩ����Z���̃Z���ʒu���s��v
' = �ˑ�    �Ȃ�
' = ����    Funcs.bas
' ==================================================================
Public Function DiffRange( _
    ByRef rTrgtRange01 As Range, _
    ByRef rTrgtRange02 As Range, _
    Optional bCellPosCheckValid As Boolean = False _
) As Boolean
    DiffRange = True
    If rTrgtRange01.Count = rTrgtRange02.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    If rTrgtRange01.Rows.Count = rTrgtRange02.Rows.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    If rTrgtRange01.Columns.Count = rTrgtRange02.Columns.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    Dim lIdx As Long
    For lIdx = 1 To rTrgtRange01.Count
        If rTrgtRange01(lIdx).Value = rTrgtRange02(lIdx).Value Then
            'Do Nothing
        Else
            DiffRange = False
            Exit Function
        End If
    Next lIdx
    If bCellPosCheckValid = True Then
        If rTrgtRange01(1).Row = rTrgtRange02(1).Row And _
           rTrgtRange01(1).Column = rTrgtRange02(1).Column And _
           rTrgtRange01(rTrgtRange01.Count).Row = rTrgtRange02(rTrgtRange02.Count).Row And _
           rTrgtRange01(rTrgtRange01.Count).Column = rTrgtRange02(rTrgtRange02.Count).Column Then
            'Do Nothing
        Else
            DiffRange = False
            Exit Function
        End If
    Else
        'Do Nothing
    End If
    
End Function
    Private Sub Test_DiffRange()
        Dim shDiff01 As Worksheet
        Dim shDiff02 As Worksheet
        Set shDiff01 = ThisWorkbook.Sheets("�^�O�ꗗ")
        Set shDiff02 = ThisWorkbook.Sheets("�^�O�ꗗ_�~���[")
        Debug.Print DiffRange( _
            shDiff01.Range( _
                shDiff01.Cells(4, 6), _
                shDiff01.Cells(4, 39) _
            ), _
            shDiff02.Range( _
                shDiff02.Cells(4, 6), _
                shDiff02.Cells(4, 39) _
            ) _
        )
    End Sub

' ==================================================================
' = �T�v    �t�@�C��/�t�H���_�̑��݊m�F���s��
' = ����    sFileDirPath  String  [in]  �t�@�C��/�t�H���_�̃p�X
' = �ߒl                  Boolean       ���݊m�F����
' = �o��    ���[�J�����̃t�@�C���p�X���w�肷�邱�ƁB
' =         URL ���w�肵���ꍇ�A�����݂Ƃ݂Ȃ����B
' =         �t�@�C���T�[�o�[�͎w��\�B
' = �ˑ�    �Ȃ�
' = ����    Funcs.bas
' ==================================================================
Public Function Exists( _
    ByVal sFileDirPath As String _
) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(sFileDirPath) Then
        Exists = True
    Else
        If objFSO.FileExists(sFileDirPath) Then
            Exists = True
        Else
            Exists = False
        End If
    End If
End Function
    Private Sub Test_Exists()
        Dim asTestPath() As String
        Dim sTestDir As String
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        sTestDir = objWshShell.SpecialFolders("Templates") & "\"
        ReDim Preserve asTestPath(4)
        asTestPath(0) = sTestDir & "is_exist_test_path01.txt"
        asTestPath(1) = sTestDir & "is_exist_test_path02.txt"
        asTestPath(2) = sTestDir & "is_exist_test_path03"
        asTestPath(3) = sTestDir & "is_exist_test_path04"
        asTestPath(4) = "https://www49.atwiki.jp/draemonash/"
        
        '�t�@�C���f�B���N�g���쐬
        Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Dim i As Long
        Open asTestPath(0) For Output As #1
        Print #1, "for Vba IsExist()'s test"
        Close #1
        objFSO.CreateFolder (asTestPath(2))
        
        '�e�X�g���s
        Dim sMsg As String
        For i = LBound(asTestPath) To UBound(asTestPath)
            sMsg = sMsg & vbNewLine & Exists(asTestPath(i))
        Next i
        Debug.Print sMsg
        
        '�t�@�C���f�B���N�g���폜
        For i = LBound(asTestPath) To UBound(asTestPath)
            If objFSO.FolderExists(asTestPath(i)) Then
                objFSO.DeleteFolder asTestPath(i), True
            Else
                If objFSO.FileExists(asTestPath(i)) Then
                    objFSO.DeleteFile asTestPath(i), True
                Else
                    'Do Nothing
                End If
            End If
        Next i
    End Sub

' ==================================================================
' = �T�v    �����񒷂���^�u����؂�ʒu�܂ł̃^�u��������ԋp����
' = ����    lLen        Long    [in]  ������
' = ����    lLenMax     Long    [in]  ������ő咷�i�ȗ��j
' = ����    lTabWidth   Long    [in]  �^�u�������i�ȗ��j
' = �ߒl                Variant       �^�u������
' = �o��    ���s��)lLen:3,lLenMax:9,lTabWidth:4
'             �uxxx^^   ^   �v
'             �uxxxxxxxxx^  �v
'               ��return:3
' = �ˑ�    Mng_String.bas/CalcPaddingWidth()
' = ����    Mng_String.bas
' ==================================================================
Public Function CalcPaddingTabWidth( _
    ByVal lLen As Long, _
    Optional ByVal lLenMax As Long = 0, _
    Optional ByVal lTabWidth As Long = 4 _
) As Variant
    Dim vPaddingWidth As Variant
    If lTabWidth = 0 Then
        CalcPaddingTabWidth = xlErrDiv0
    ElseIf lTabWidth < 0 Or lLen < 0 Or lLenMax < 0 Then
        CalcPaddingTabWidth = xlErrValue
    Else
        '�p�f�B���O��(�X�y�[�X)�Z�o
        vPaddingWidth = CalcPaddingWidth(lLen, lLenMax, lTabWidth)
        '�p�f�B���O��(�^�u)�Z�o
        CalcPaddingTabWidth = Application.WorksheetFunction.RoundUp(vPaddingWidth / lTabWidth, 0)
    End If
End Function
    Private Sub Test_CalcPaddingTabWidth()
        Debug.Print "*** test start! ***"
        Debug.Print CalcPaddingTabWidth(2, 0, 4) = 1
        Debug.Print CalcPaddingTabWidth(2, 1, 4) = 1
        Debug.Print CalcPaddingTabWidth(2, 4, 4) = 2
        Debug.Print CalcPaddingTabWidth(2, 6, 4) = 2
        Debug.Print CalcPaddingTabWidth(4, 0, 4) = 1
        Debug.Print CalcPaddingTabWidth(0, 0, 4) = 1
        Debug.Print CalcPaddingTabWidth(0, 2, 4) = 1
        Debug.Print CalcPaddingTabWidth(0, 3, 4) = 1
        Debug.Print CalcPaddingTabWidth(0, 4, 4) = 2
        Debug.Print CalcPaddingTabWidth(0, 5, 4) = 2
        
        Debug.Print CalcPaddingTabWidth(5, 19, 4) = 4
        Debug.Print CalcPaddingTabWidth(5, 20, 4) = 5
        Debug.Print CalcPaddingTabWidth(5, 21, 4) = 5
        Debug.Print CalcPaddingTabWidth(5, 22, 4) = 5
        Debug.Print CalcPaddingTabWidth(5, 23, 4) = 5
        Debug.Print CalcPaddingTabWidth(5, 24, 4) = 6
        
        Debug.Print CalcPaddingTabWidth(5, 19) = 4
        Debug.Print CalcPaddingTabWidth(5, 20) = 5
        Debug.Print CalcPaddingTabWidth(5, 21) = 5
        Debug.Print CalcPaddingTabWidth(5, 22) = 5
        Debug.Print CalcPaddingTabWidth(5, 23) = 5
        Debug.Print CalcPaddingTabWidth(5, 24) = 6
        
        Debug.Print CalcPaddingTabWidth(0) = 1
        Debug.Print CalcPaddingTabWidth(3) = 1
        Debug.Print CalcPaddingTabWidth(4) = 1
        Debug.Print CalcPaddingTabWidth(5) = 1
        Debug.Print CalcPaddingTabWidth(6) = 1
        
        Debug.Print CalcPaddingTabWidth(5, 15, 8) = 2
        Debug.Print CalcPaddingTabWidth(5, 16, 8) = 3
        Debug.Print CalcPaddingTabWidth(5, 17, 8) = 3
        Debug.Print CalcPaddingTabWidth(5, 18, 8) = 3
        Debug.Print CalcPaddingTabWidth(5, 19, 8) = 3
        Debug.Print CalcPaddingTabWidth(5, 20, 8) = 3
        Debug.Print CalcPaddingTabWidth(5, 21, 8) = 3
        Debug.Print CalcPaddingTabWidth(5, 22, 8) = 3
        Debug.Print CalcPaddingTabWidth(5, 23, 8) = 3
        Debug.Print CalcPaddingTabWidth(5, 24, 8) = 4
        
        Debug.Print CalcPaddingTabWidth(1, 5, 0) = xlErrDiv0
        Debug.Print CalcPaddingTabWidth(1, 5, -1) = xlErrValue
        Debug.Print CalcPaddingTabWidth(1, -1, 4) = xlErrValue
        Debug.Print CalcPaddingTabWidth(-1, 5, 4) = xlErrValue
        
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �����񒷂����؂蕝�ʒu�܂ł̕�������ԋp����
' = ����    lLen        Long    [in]  ������
' = ����    lLenMax     Long    [in]  ������ő咷�i�ȗ��j
' = ����    lSepWidth   Long    [in]  ��؂蕝�i�ȗ��j
' = �ߒl                Variant       ������
' = �o��    ���s��)lLen:3,lLenMax:9,lSepWidth:4
'             �uxxx         �v
'             �uxxxxxxxxx   �v
'               ��return:9
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function CalcPaddingWidth( _
    ByVal lLen As Long, _
    Optional ByVal lLenMax As Long = 0, _
    Optional ByVal lSepWidth As Long = 4 _
) As Variant
    Dim lPaddingWidth As Long
    If lSepWidth = 0 Then
        CalcPaddingWidth = xlErrDiv0
    ElseIf lSepWidth < 0 Or lLen < 0 Or lLenMax < 0 Then
        CalcPaddingWidth = xlErrValue
    Else
        If lLen > lLenMax Then
            lLenMax = lLen
        End If
        lPaddingWidth = lSepWidth - (lLenMax Mod lSepWidth)
        CalcPaddingWidth = (lPaddingWidth + lLenMax) - lLen
    End If
End Function
    Private Sub Test_CalcPaddingWidth()
        Debug.Print "*** test start! ***"
        Debug.Print CalcPaddingWidth(0, 5, 4) = 8
        Debug.Print CalcPaddingWidth(3, 5, 4) = 5
        Debug.Print CalcPaddingWidth(4, 5, 4) = 4
        Debug.Print CalcPaddingWidth(5, 5, 4) = 3
        Debug.Print CalcPaddingWidth(6, 5, 4) = 2
        Debug.Print CalcPaddingWidth(7, 5, 4) = 1
        Debug.Print CalcPaddingWidth(8, 5, 4) = 4
        
        Debug.Print CalcPaddingWidth(0, 1, 8) = 8
        Debug.Print CalcPaddingWidth(1, 1, 8) = 7
        Debug.Print CalcPaddingWidth(2, 1, 8) = 6
        Debug.Print CalcPaddingWidth(3, 1, 8) = 5
        Debug.Print CalcPaddingWidth(4, 1, 8) = 4
        Debug.Print CalcPaddingWidth(5, 1, 8) = 3
        Debug.Print CalcPaddingWidth(6, 1, 8) = 2
        Debug.Print CalcPaddingWidth(7, 1, 8) = 1
        Debug.Print CalcPaddingWidth(8, 1, 8) = 8
        
        Debug.Print CalcPaddingWidth(0, 5, 7) = 7
        Debug.Print CalcPaddingWidth(3, 5, 7) = 4
        Debug.Print CalcPaddingWidth(4, 5, 7) = 3
        Debug.Print CalcPaddingWidth(5, 5, 7) = 2
        Debug.Print CalcPaddingWidth(6, 5, 7) = 1
        Debug.Print CalcPaddingWidth(7, 5, 7) = 7
        Debug.Print CalcPaddingWidth(8, 5, 7) = 6
        
        Debug.Print CalcPaddingWidth(1, 5, 0) = xlErrDiv0
        Debug.Print CalcPaddingWidth(1, 5, -1) = xlErrValue
        Debug.Print CalcPaddingWidth(1, -1, 4) = xlErrValue
        Debug.Print CalcPaddingWidth(-1, 5, 4) = xlErrValue
        
        Debug.Print CalcPaddingWidth(0, 5) = 8
        Debug.Print CalcPaddingWidth(3, 5) = 5
        Debug.Print CalcPaddingWidth(4, 5) = 4
        Debug.Print CalcPaddingWidth(5, 5) = 3
        Debug.Print CalcPaddingWidth(6, 5) = 2
        Debug.Print CalcPaddingWidth(7, 5) = 1
        Debug.Print CalcPaddingWidth(8, 5) = 4
        
        Debug.Print CalcPaddingWidth(0) = 4
        Debug.Print CalcPaddingWidth(3) = 1
        Debug.Print CalcPaddingWidth(4) = 4
        Debug.Print CalcPaddingWidth(5) = 3
        Debug.Print CalcPaddingWidth(6) = 2
        Debug.Print CalcPaddingWidth(7) = 1
        Debug.Print CalcPaddingWidth(8) = 4

        Debug.Print "*** test finished! ***"
    End Sub

'********************************************************************************
'* �����֐���`
'********************************************************************************
' ==================================================================
' = �T�v    �]�艉�Z
' =         Mod ���Z�q�� 2,147,483,647 ���傫�������̓I�[�o�[�t���[����B
' =         �{�֐��͏�L�ȏ�̐��l���������Ƃ��ł���B
' = ����    cNum1   Currency    [in]    �l1
' = ����    cNum2   Currency    [in]    �l2
' = �ߒl            Currency            ���Z����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Math.bas
' ==================================================================
Private Function ModEx( _
    ByVal cNum1 As Currency, _
    ByVal cNum2 As Currency _
) As Currency
    ModEx = CDec(cNum1) - Fix(CDec(cNum1) / CDec(cNum2)) * CDec(cNum2)
End Function
    Private Sub Test_ModEx()
        Debug.Print "*** test start! ***"
        Debug.Print ModEx(12, 2)             '0
        Debug.Print ModEx(12, 3)             '0
        Debug.Print ModEx(12, 5)             '2
        Debug.Print ModEx(2147483647, 5)     '2
        Debug.Print ModEx(2147483648#, 5)    '3
        Debug.Print ModEx(2147483649#, 5)    '4
        Debug.Print ModEx(-2147483647, 5)    '-2
        Debug.Print ModEx(5, 2147483648#)    '5
        Debug.Print ModEx(5, 2147483649#)    '5
        Debug.Print ModEx(0, 5)              '0
       'Debug.Print ModEx(5, 0)              '�v���O������~
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    ��ϊ� 10�i����16�i��(32bit�p)
' = ����    cInDecVal   Currency    [in]    10�i��
' = �ߒl                String              �ϊ�����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Private Function Dec2Hex( _
    ByVal cInDecVal As Currency _
) As String
    Dim cInDecValHi As Currency
    Dim cInDecValLo As Currency
    Dim sOutHexValHi As String
    Dim sOutHexValLo As String
    Dim sHexVal As String
    If cInDecVal < -2147483648# Or cInDecVal > 4294967295# Then
        Dec2Hex = "error"
        Exit Function
    End If
    
    If cInDecVal >= 0 Then
        'Do Nothing
    Else
        cInDecVal = cInDecVal + 4294967296#
    End If
    cInDecVal = Int(ModEx(cInDecVal, 4294967296#))
    
    cInDecValHi = Int(cInDecVal / 65536)
    cInDecValLo = Int(ModEx(cInDecVal, 65536))
    sOutHexValHi = UCase(String(4 - Len(Hex(cInDecValHi)), "0") & Hex(cInDecValHi))
    sOutHexValLo = UCase(String(4 - Len(Hex(cInDecValLo)), "0") & Hex(cInDecValLo))
    Dec2Hex = sOutHexValHi & sOutHexValLo
End Function
    Private Sub Test_Dec2Hex()
        Debug.Print "*** test start! ***"
        Debug.Print Dec2Hex(0)              '00000000
        Debug.Print Dec2Hex(1)              '00000001
        Debug.Print Dec2Hex(2)              '00000002
        Debug.Print Dec2Hex(10)             '0000000A
        Debug.Print Dec2Hex(15)             '0000000F
        Debug.Print Dec2Hex(16)             '00000010
        Debug.Print Dec2Hex(4294967296#)    'error
        Debug.Print Dec2Hex(4294967295#)    'FFFFFFFF
        Debug.Print Dec2Hex(4294967294#)    'FFFFFFFE
        Debug.Print Dec2Hex(2147483648#)    '80000000
        Debug.Print Dec2Hex(2147483647)     '7FFFFFFF
        Debug.Print Dec2Hex(2147483646)     '7FFFFFFE
        Debug.Print Dec2Hex(65536)          '00010000
        Debug.Print Dec2Hex(65535)          '0000FFFF
        Debug.Print Dec2Hex(65534)          '0000FFFE
        Debug.Print Dec2Hex(2)              '00000002
        Debug.Print Dec2Hex(1)              '00000001
        Debug.Print Dec2Hex(0)              '00000000
        Debug.Print Dec2Hex(-1)             'FFFFFFFF
        Debug.Print Dec2Hex(-2)             'FFFFFFFE
        Debug.Print Dec2Hex(-65534)         'FFFF0002
        Debug.Print Dec2Hex(-65535)         'FFFF0001
        Debug.Print Dec2Hex(-65536)         'FFFF0000
        Debug.Print Dec2Hex(-2147483647)    '80000001
        Debug.Print Dec2Hex(-2147483648#)   '80000000
        Debug.Print Dec2Hex(-2147483649#)   'error
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    ��ϊ� 16�i����10�i��(32bit�p)
' = ����    sInHexVal       String      [in]    16�i��
' = ����    bIsSignEnable   Boolean     [in]    �����L��
' = �ߒl                    Variant             �ϊ�����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Private Function Hex2Dec( _
    ByVal sInHexVal As String, _
    ByVal bIsSignEnable As Boolean _
) As Variant
    If Len(sInHexVal) <> 8 Then
        Hex2Dec = "error"
        Exit Function
    End If
    Dim cInDecValHi As Currency
    Dim cInDecValLo As Currency
    Dim cOutDecVal As Currency
    On Error Resume Next
    cInDecValHi = CCur("&H" & Left$(sInHexVal, 4)) * 65536
    cInDecValLo = CCur("&H" & Right$(sInHexVal, 4))
    cOutDecVal = cInDecValHi + cInDecValLo
    If Err.Number <> 0 Then
        Hex2Dec = "error"
        Err.Clear
    Else
        If bIsSignEnable = True Then
            If cOutDecVal > 2147483647 Then
                Hex2Dec = cOutDecVal - 4294967296#
            Else
                Hex2Dec = cOutDecVal
            End If
        Else
            Hex2Dec = cOutDecVal
        End If
    End If
    On Error GoTo 0
End Function
    Private Sub Test_Hex2Dec()
        Debug.Print "*** test start! ***"
        Debug.Print Hex2Dec("00000000", False) '0
        Debug.Print Hex2Dec("00000001", False) '1
        Debug.Print Hex2Dec("00000009", False) '9
        Debug.Print Hex2Dec("0000000A", False) '10
        Debug.Print "<<sign>>"
        Debug.Print Hex2Dec("7FFFFFFF", True)  '2147483647
        Debug.Print Hex2Dec("7FFFFFFE", True)  '2147483646
        Debug.Print Hex2Dec("00000002", True)  '2
        Debug.Print Hex2Dec("00000001", True)  '1
        Debug.Print Hex2Dec("00000000", True)  '0
        Debug.Print Hex2Dec("FFFFFFFF", True)  '-1
        Debug.Print Hex2Dec("FFFFFFFE", True)  '-2
        Debug.Print Hex2Dec("80000001", True)  '-2147483647
        Debug.Print Hex2Dec("80000000", True)  '-2147483648
        Debug.Print "<<unsign>>"
        Debug.Print Hex2Dec("FFFFFFFF", False) '4294967295
        Debug.Print Hex2Dec("FFFFFFFE", False) '4294967294
        Debug.Print Hex2Dec("80000001", False) '2147483649
        Debug.Print Hex2Dec("80000000", False) '2147483648
        Debug.Print Hex2Dec("00000001", False) '1
        Debug.Print Hex2Dec("00000000", False) '0
        Debug.Print Hex2Dec("0000000", False)  'error
        Debug.Print Hex2Dec("000000000", False) 'error
        Debug.Print Hex2Dec("8000000K", False) 'error
        Debug.Print Hex2Dec("80 00001", False) 'error
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    ��ϊ� 16�i����10�i��
' = ����    sHexVal         String      [in]    16�i��
' = �ߒl                    String              �ϊ�����
' = �o��    �w��͈͈ȊO�̒l���w�肷��ƕ����� "error" ��ԋp����B
' = �ˑ�    Mng_String.bas/Hex2BinMap()
' = ����    Mng_String.bas
' ==================================================================
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

' ==================================================================
' = �T�v    ��ϊ� 2�i����16�i��
' = ����    sBinVal         String      [in]    2�i��
' = ����    bIsUcase        Boolean     [in]    �啶��������
' = �ߒl                    String              �ϊ�����
' = �o��    �w��͈͈ȊO�̒l���w�肷��ƕ����� "error" ��ԋp����B
' = �ˑ�    Mng_String.bas/Bin2HexMap()
' = ����    Mng_String.bas
' ==================================================================
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

' ==================================================================
' = �T�v    ��ϊ��p�}�b�v 16�i����2�i��
' = ����    sHexVal         String      [in]    16�i��
' = �ߒl                    String              �ϊ�����
' = �o��    �w��͈͈ȊO�̒l���w�肷��ƕ����� "error" ��ԋp����B
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
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

' ==================================================================
' = �T�v    ��ϊ��p�}�b�v 2�i����16�i��
' = ����    sBinVal         String      [in]    2�i��
' = ����    bIsUcase        Boolean     [in]    �啶��������
' = �ߒl                    String              �ϊ�����
' = �o��    �w��͈͈ȊO�̒l���w�肷��ƕ����� "error" ��ԋp����B
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
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

' ==================================================================
' = �T�v    �_���r�b�g�V�t�g�i������Łj
' = ����    sInBinVal       String              [in]    2�i��
' = ����    lInShiftNum     Long                [in]    �V�t�g��
' = ����    eInDirection    E_SHIFT_DIRECTiON   [in]    �V�t�g����
' = ����    lInBitLen       Long                [in]    �r�b�g�T�C�Y
' = �ߒl                    String                      �V�t�g����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Private Function BitShiftLogStrBin( _
    ByVal sInBinVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    ByVal lInBitLen As Long _
) As String
    Debug.Assert sInBinVal <> ""
    Debug.Assert Replace(Replace(sInBinVal, "1", ""), "0", "") = ""
    Debug.Assert lInShiftNum >= 0
    Debug.Assert eInDirection = LEFT_SHIFT Or RIGHT_SHIFT
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
        Debug.Print BitShiftLogStrBin("0", 0, LEFT_SHIFT, 8)       '00000000
        Debug.Print BitShiftLogStrBin("0", 2, LEFT_SHIFT, 8)       '00000000
        Debug.Print BitShiftLogStrBin("1", 0, LEFT_SHIFT, 8)       '00000001
        Debug.Print BitShiftLogStrBin("1", 2, LEFT_SHIFT, 8)       '00000100
        Debug.Print BitShiftLogStrBin("1", 7, LEFT_SHIFT, 8)       '10000000
        Debug.Print BitShiftLogStrBin("1", 8, LEFT_SHIFT, 8)       '00000000
        Debug.Print BitShiftLogStrBin("1", 0, RIGHT_SHIFT, 8)      '00000001
        Debug.Print BitShiftLogStrBin("1", 1, RIGHT_SHIFT, 8)      '00000000
        Debug.Print BitShiftLogStrBin("1", 2, RIGHT_SHIFT, 8)      '00000000
        Debug.Print BitShiftLogStrBin("1011", 0, LEFT_SHIFT, 0)    '1011
        Debug.Print BitShiftLogStrBin("1011", 1, LEFT_SHIFT, 0)    '10110
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 0)    '101100
        Debug.Print BitShiftLogStrBin("1011", 0, RIGHT_SHIFT, 0)   '1011
        Debug.Print BitShiftLogStrBin("1011", 1, RIGHT_SHIFT, 0)   '101
        Debug.Print BitShiftLogStrBin("1011", 2, RIGHT_SHIFT, 0)   '10
        Debug.Print BitShiftLogStrBin("1011", 3, RIGHT_SHIFT, 0)   '1
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, 0)   '0
        Debug.Print BitShiftLogStrBin("1011", 5, RIGHT_SHIFT, 0)   '0
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 2)    '00
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 3)    '100
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 4)    '1100
        Debug.Print BitShiftLogStrBin("1011", 2, LEFT_SHIFT, 10)   '0000101100
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, 8)   '00000000
        Debug.Print BitShiftLogStrBin("1011", 4, RIGHT_SHIFT, 8)   '00000000
       'Debug.Print BitShiftLogStrBin("", 2, LEFT_SHIFT,  10)       '�v���O������~
       'Debug.Print BitShiftLogStrBin("101A", 1, LEFT_SHIFT,  10)   '�v���O������~
       'Debug.Print BitShiftLogStrBin("1011", -1, LEFT_SHIFT,  10)  '�v���O������~
       'Debug.Print BitShiftLogStrBin("1011", 1, 5,  10)            '�v���O������~
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �Z�p�r�b�g�V�t�g�i������Łj
' = ����    sInBinVal           String              [in]    2�i��
' = ����    lInShiftNum         Long                [in]    �V�t�g��
' = ����    eInDirection        E_SHIFT_DIRECTiON   [in]    �V�t�g����
' = ����    lInBitLen           Long                [in]    �r�b�g�T�C�Y
' = ����    bIsSaveSignBit      Boolean             [in]    �ȉ��A�Q��
' = ����    bIsExecAutoAlign    Boolean             [in]    �ȉ��A�Q��
' = �ߒl                        String                      �V�t�g����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Private Function BitShiftAriStrBin( _
    ByVal sInBinVal As String, _
    ByVal lInShiftNum As Long, _
    ByVal eInDirection As E_SHIFT_DIRECTiON, _
    ByVal lInBitLen As Long, _
    ByVal bIsSaveSignBit As Boolean, _
    ByVal bIsExecAutoAlign As Boolean _
) As String
    ' <<bIsSaveSignBit>>
    '   �o�͌��������͒l�i������j�̒��������������ꍇ�ɁA
    '   �����r�b�g��ێ����邩�A�������Đ؂�̂Ă邩��I������B
    '     True  : �����r�b�g��ێ�����
    '               ex1) "AB"(0b10101011) ���o�͌���1�Ƃ��ĉE1�Z�p(�����r�b�g�ێ�)�V�t�g
    '                 �� "D"(0b1101)
    '     False : �����r�b�g��؂�̂Ă�
    '               ex2) "AB"(0b10101011) ���o�͌���1�Ƃ��ĉE1�Z�p(�����r�b�g�؎�)�V�t�g
    '                 �� "5"(0b0101)
    ' <<bIsExecAutoAlign>>
    '   �o�͌��ʂ�8�r�b�g���E�ɑ����邩�ǂ�����I������B
    '     True  : ������
    '               ex1) 10101011 ���E1�r�b�g�V�t�g
    '                 �� 11010101
    '               ex2) 10101011 ����1�r�b�g�V�t�g
    '                 �� 1111111101010110
    '     False : �����Ȃ�
    '               ex1) 10101011 ���E1�r�b�g�V�t�g
    '                 ��  1010101
    '               ex2) 10101011 ����1�r�b�g�V�t�g
    '                 �� 101010110
    
    Debug.Assert sInBinVal <> ""
    Debug.Assert Replace(Replace(sInBinVal, "1", ""), "0", "") = ""
    Debug.Assert lInShiftNum >= 0
    Debug.Assert eInDirection = LEFT_SHIFT Or RIGHT_SHIFT
    Debug.Assert lInBitLen >= 0
    If bIsExecAutoAlign = True Then
        Debug.Assert Len(sInBinVal) = 8
        Debug.Assert lInBitLen Mod 8 = 0
    Else
        'Do Nothing
    End If
    
    '�r�b�g�V�t�g
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
    If lInBitLen = 0 Then
        If bIsExecAutoAlign = True Then
            Dim sPadBit As String
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
            If bIsSaveSignBit = True Then
                BitShiftAriStrBin = sInSignBit & Right$(sOutLogicBit, lInBitLen - 1)
            Else
                BitShiftAriStrBin = Right$(sTmpBinVal, lInBitLen)
            End If
        Else
            BitShiftAriStrBin = sTmpBinVal
        End If
    End If
End Function
    Private Sub Test_BitShiftAriStrBin()
        Debug.Print "*** test start! ***"
        Debug.Print "<<test 001-01>>"                                                   '<<test 001-01>>
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 0, True, False)       '01001011
        Debug.Print BitShiftAriStrBin("01001011", 1, RIGHT_SHIFT, 0, True, False)       '0100101
        Debug.Print BitShiftAriStrBin("01001011", 2, RIGHT_SHIFT, 0, True, False)       '010010
        Debug.Print BitShiftAriStrBin("01001011", 3, RIGHT_SHIFT, 0, True, False)       '01001
        Debug.Print BitShiftAriStrBin("01001011", 4, RIGHT_SHIFT, 0, True, False)       '0100
        Debug.Print BitShiftAriStrBin("01001011", 5, RIGHT_SHIFT, 0, True, False)       '010
        Debug.Print BitShiftAriStrBin("01001011", 6, RIGHT_SHIFT, 0, True, False)       '01
        Debug.Print BitShiftAriStrBin("01001011", 7, RIGHT_SHIFT, 0, True, False)       '0
        Debug.Print BitShiftAriStrBin("01001011", 8, RIGHT_SHIFT, 0, True, False)       '0
        Debug.Print BitShiftAriStrBin("01001011", 9, RIGHT_SHIFT, 0, True, False)       '0
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 10, True, False)      '0001001011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 5, True, False)       '01011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 4, True, False)       '0011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 3, True, False)       '011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 2, True, False)       '01
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 1, True, False)       '0
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 0, True, False)        '01001011
        Debug.Print BitShiftAriStrBin("01001011", 1, LEFT_SHIFT, 0, True, False)        '010010110
        Debug.Print BitShiftAriStrBin("01001011", 2, LEFT_SHIFT, 0, True, False)        '0100101100
        Debug.Print BitShiftAriStrBin("01001011", 5, LEFT_SHIFT, 0, True, False)        '0100101100000
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 10, True, False)       '0001001011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 5, True, False)        '01011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 4, True, False)        '0011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 3, True, False)        '011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 2, True, False)        '01
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 1, True, False)        '0
        Debug.Print "<<test 001-02>>"                                                   '<<test 001-02>>
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 0, False, False)      '01001011
        Debug.Print BitShiftAriStrBin("01001011", 1, RIGHT_SHIFT, 0, False, False)      '0100101
        Debug.Print BitShiftAriStrBin("01001011", 2, RIGHT_SHIFT, 0, False, False)      '010010
        Debug.Print BitShiftAriStrBin("01001011", 3, RIGHT_SHIFT, 0, False, False)      '01001
        Debug.Print BitShiftAriStrBin("01001011", 4, RIGHT_SHIFT, 0, False, False)      '0100
        Debug.Print BitShiftAriStrBin("01001011", 5, RIGHT_SHIFT, 0, False, False)      '010
        Debug.Print BitShiftAriStrBin("01001011", 6, RIGHT_SHIFT, 0, False, False)      '01
        Debug.Print BitShiftAriStrBin("01001011", 7, RIGHT_SHIFT, 0, False, False)      '0
        Debug.Print BitShiftAriStrBin("01001011", 8, RIGHT_SHIFT, 0, False, False)      '0
        Debug.Print BitShiftAriStrBin("01001011", 9, RIGHT_SHIFT, 0, False, False)      '0
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 10, False, False)     '0001001011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 5, False, False)      '01011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 4, False, False)      '1011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 3, False, False)      '011
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 2, False, False)      '11
        Debug.Print BitShiftAriStrBin("01001011", 0, RIGHT_SHIFT, 1, False, False)      '1
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 0, False, False)       '01001011
        Debug.Print BitShiftAriStrBin("01001011", 1, LEFT_SHIFT, 0, False, False)       '010010110
        Debug.Print BitShiftAriStrBin("01001011", 2, LEFT_SHIFT, 0, False, False)       '0100101100
        Debug.Print BitShiftAriStrBin("01001011", 5, LEFT_SHIFT, 0, False, False)       '0100101100000
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 10, False, False)      '0001001011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 5, False, False)       '01011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 4, False, False)       '1011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 3, False, False)       '011
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 2, False, False)       '11
        Debug.Print BitShiftAriStrBin("01001011", 0, LEFT_SHIFT, 1, False, False)       '1
        Debug.Print "<<test 001-03>>"                                                   '<<test 001-03>>
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 0, True, False)       '10001011
        Debug.Print BitShiftAriStrBin("10001011", 1, RIGHT_SHIFT, 0, True, False)       '1000101
        Debug.Print BitShiftAriStrBin("10001011", 2, RIGHT_SHIFT, 0, True, False)       '100010
        Debug.Print BitShiftAriStrBin("10001011", 3, RIGHT_SHIFT, 0, True, False)       '10001
        Debug.Print BitShiftAriStrBin("10001011", 4, RIGHT_SHIFT, 0, True, False)       '1000
        Debug.Print BitShiftAriStrBin("10001011", 5, RIGHT_SHIFT, 0, True, False)       '100
        Debug.Print BitShiftAriStrBin("10001011", 6, RIGHT_SHIFT, 0, True, False)       '10
        Debug.Print BitShiftAriStrBin("10001011", 7, RIGHT_SHIFT, 0, True, False)       '1
        Debug.Print BitShiftAriStrBin("10001011", 8, RIGHT_SHIFT, 0, True, False)       '1
        Debug.Print BitShiftAriStrBin("10001011", 9, RIGHT_SHIFT, 0, True, False)       '1
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 10, True, False)      '1110001011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 5, True, False)       '11011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 4, True, False)       '1011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 3, True, False)       '111
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 2, True, False)       '11
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 1, True, False)       '1
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 0, True, False)        '10001011
        Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, 0, True, False)        '100010110
        Debug.Print BitShiftAriStrBin("10001011", 2, LEFT_SHIFT, 0, True, False)        '1000101100
        Debug.Print BitShiftAriStrBin("10001011", 5, LEFT_SHIFT, 0, True, False)        '1000101100000
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 10, True, False)       '1110001011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 5, True, False)        '11011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 4, True, False)        '1011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 3, True, False)        '111
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 2, True, False)        '11
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 1, True, False)        '1
        Debug.Print "<<test 001-04>>"                                                   '<<test 001-04>>
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 0, False, False)      '10001011
        Debug.Print BitShiftAriStrBin("10001011", 1, RIGHT_SHIFT, 0, False, False)      '1000101
        Debug.Print BitShiftAriStrBin("10001011", 2, RIGHT_SHIFT, 0, False, False)      '100010
        Debug.Print BitShiftAriStrBin("10001011", 3, RIGHT_SHIFT, 0, False, False)      '10001
        Debug.Print BitShiftAriStrBin("10001011", 4, RIGHT_SHIFT, 0, False, False)      '1000
        Debug.Print BitShiftAriStrBin("10001011", 5, RIGHT_SHIFT, 0, False, False)      '100
        Debug.Print BitShiftAriStrBin("10001011", 6, RIGHT_SHIFT, 0, False, False)      '10
        Debug.Print BitShiftAriStrBin("10001011", 7, RIGHT_SHIFT, 0, False, False)      '1
        Debug.Print BitShiftAriStrBin("10001011", 8, RIGHT_SHIFT, 0, False, False)      '1
        Debug.Print BitShiftAriStrBin("10001011", 9, RIGHT_SHIFT, 0, False, False)      '1
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 10, False, False)     '1110001011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 5, False, False)      '01011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 4, False, False)      '1011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 3, False, False)      '011
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 2, False, False)      '11
        Debug.Print BitShiftAriStrBin("10001011", 0, RIGHT_SHIFT, 1, False, False)      '1
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 0, False, False)       '10001011
        Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, 0, False, False)       '100010110
        Debug.Print BitShiftAriStrBin("10001011", 2, LEFT_SHIFT, 0, False, False)       '1000101100
        Debug.Print BitShiftAriStrBin("10001011", 5, LEFT_SHIFT, 0, False, False)       '1000101100000
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 10, False, False)      '1110001011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 5, False, False)       '01011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 4, False, False)       '1011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 3, False, False)       '011
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 2, False, False)       '11
        Debug.Print BitShiftAriStrBin("10001011", 0, LEFT_SHIFT, 1, False, False)       '1
    '   Debug.Print "<<test 001-05>>"                                                   '<<test 001-05>>
    '   Debug.Print BitShiftAriStrBin("10001011", 1, 5, 10, True, False)                '�v���O������~
    '   Debug.Print BitShiftAriStrBin("", 2, LEFT_SHIFT, 10, True, False)               '�v���O������~
    '   Debug.Print BitShiftAriStrBin("10001011", -1, LEFT_SHIFT, 10, True, False)      '�v���O������~
    '   Debug.Print BitShiftAriStrBin("1000101A", 1, LEFT_SHIFT, 10, True, False)       '�v���O������~
    '   Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, -1, True, False)       '�v���O������~
        Debug.Print "<<test 002-01>>"                                                   '<<test 002-01>>"
        Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, 0, True, True)        '00101011
        Debug.Print BitShiftAriStrBin("00101011", 1, RIGHT_SHIFT, 0, True, True)        '00010101
        Debug.Print BitShiftAriStrBin("00101011", 2, RIGHT_SHIFT, 0, True, True)        '00001010
        Debug.Print BitShiftAriStrBin("00101011", 3, RIGHT_SHIFT, 0, True, True)        '00000101
        Debug.Print BitShiftAriStrBin("00101011", 4, RIGHT_SHIFT, 0, True, True)        '00000010
        Debug.Print BitShiftAriStrBin("00101011", 5, RIGHT_SHIFT, 0, True, True)        '00000001
        Debug.Print BitShiftAriStrBin("00101011", 6, RIGHT_SHIFT, 0, True, True)        '00000000
        Debug.Print BitShiftAriStrBin("00101011", 7, RIGHT_SHIFT, 0, True, True)        '00000000
        Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, 16, True, True)       '0000000000101011
        Debug.Print BitShiftAriStrBin("00101011", 0, RIGHT_SHIFT, 8, True, True)        '00101011
        Debug.Print BitShiftAriStrBin("00101011", 1, RIGHT_SHIFT, 8, True, True)        '00010101
        Debug.Print BitShiftAriStrBin("00101011", 2, RIGHT_SHIFT, 8, True, True)        '00001010
        Debug.Print BitShiftAriStrBin("00101011", 3, RIGHT_SHIFT, 8, True, True)        '00000101
        Debug.Print BitShiftAriStrBin("00101011", 4, RIGHT_SHIFT, 8, True, True)        '00000010
        Debug.Print BitShiftAriStrBin("00101011", 5, RIGHT_SHIFT, 8, True, True)        '00000001
        Debug.Print BitShiftAriStrBin("00101011", 6, RIGHT_SHIFT, 8, True, True)        '00000000
        Debug.Print BitShiftAriStrBin("00101011", 7, RIGHT_SHIFT, 8, True, True)        '00000000
        Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, 0, True, True)         '00101011
        Debug.Print BitShiftAriStrBin("00101011", 1, LEFT_SHIFT, 0, True, True)         '0000000001010110
        Debug.Print BitShiftAriStrBin("00101011", 2, LEFT_SHIFT, 0, True, True)         '0000000010101100
        Debug.Print BitShiftAriStrBin("00101011", 5, LEFT_SHIFT, 0, True, True)         '0000010101100000
        Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, 16, True, True)        '0000000000101011
        Debug.Print BitShiftAriStrBin("00101011", 0, LEFT_SHIFT, 8, True, True)         '00101011
        Debug.Print BitShiftAriStrBin("00101011", 1, LEFT_SHIFT, 8, True, True)         '01010110
        Debug.Print BitShiftAriStrBin("00101011", 2, LEFT_SHIFT, 8, True, True)         '00101100
        Debug.Print BitShiftAriStrBin("00101011", 3, LEFT_SHIFT, 8, True, True)         '01011000
        Debug.Print BitShiftAriStrBin("00101011", 4, LEFT_SHIFT, 8, True, True)         '00110000
        Debug.Print BitShiftAriStrBin("00101011", 5, LEFT_SHIFT, 8, True, True)         '01100000
        Debug.Print BitShiftAriStrBin("00101011", 6, LEFT_SHIFT, 8, True, True)         '01000000
        Debug.Print BitShiftAriStrBin("00101011", 7, LEFT_SHIFT, 8, True, True)         '00000000
        Debug.Print "<<test 002-02>>"                                                   '<<test 002-02>>
        Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, 0, True, True)        '10101011
        Debug.Print BitShiftAriStrBin("10101011", 1, RIGHT_SHIFT, 0, True, True)        '11010101
        Debug.Print BitShiftAriStrBin("10101011", 2, RIGHT_SHIFT, 0, True, True)        '11101010
        Debug.Print BitShiftAriStrBin("10101011", 3, RIGHT_SHIFT, 0, True, True)        '11110101
        Debug.Print BitShiftAriStrBin("10101011", 4, RIGHT_SHIFT, 0, True, True)        '11111010
        Debug.Print BitShiftAriStrBin("10101011", 5, RIGHT_SHIFT, 0, True, True)        '11111101
        Debug.Print BitShiftAriStrBin("10101011", 6, RIGHT_SHIFT, 0, True, True)        '11111110
        Debug.Print BitShiftAriStrBin("10101011", 7, RIGHT_SHIFT, 0, True, True)        '11111111
        Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, 16, True, True)       '1111111110101011
        Debug.Print BitShiftAriStrBin("10101011", 0, RIGHT_SHIFT, 8, True, True)        '10101011
        Debug.Print BitShiftAriStrBin("10101011", 1, RIGHT_SHIFT, 8, True, True)        '11010101
        Debug.Print BitShiftAriStrBin("10101011", 2, RIGHT_SHIFT, 8, True, True)        '11101010
        Debug.Print BitShiftAriStrBin("10101011", 3, RIGHT_SHIFT, 8, True, True)        '11110101
        Debug.Print BitShiftAriStrBin("10101011", 4, RIGHT_SHIFT, 8, True, True)        '11111010
        Debug.Print BitShiftAriStrBin("10101011", 5, RIGHT_SHIFT, 8, True, True)        '11111101
        Debug.Print BitShiftAriStrBin("10101011", 6, RIGHT_SHIFT, 8, True, True)        '11111110
        Debug.Print BitShiftAriStrBin("10101011", 7, RIGHT_SHIFT, 8, True, True)        '11111111
        Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, 0, True, True)         '10101011
        Debug.Print BitShiftAriStrBin("10101011", 1, LEFT_SHIFT, 0, True, True)         '1111111101010110
        Debug.Print BitShiftAriStrBin("10101011", 2, LEFT_SHIFT, 0, True, True)         '1111111010101100
        Debug.Print BitShiftAriStrBin("10101011", 5, LEFT_SHIFT, 0, True, True)         '1111010101100000
        Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, 16, True, True)        '1111111110101011
        Debug.Print BitShiftAriStrBin("10101011", 0, LEFT_SHIFT, 8, True, True)         '10101011
        Debug.Print BitShiftAriStrBin("10101011", 1, LEFT_SHIFT, 8, True, True)         '11010110
        Debug.Print BitShiftAriStrBin("10101011", 2, LEFT_SHIFT, 8, True, True)         '10101100
        Debug.Print BitShiftAriStrBin("10101011", 3, LEFT_SHIFT, 8, True, True)         '11011000
        Debug.Print BitShiftAriStrBin("10101011", 4, LEFT_SHIFT, 8, True, True)         '10110000
        Debug.Print BitShiftAriStrBin("10101011", 5, LEFT_SHIFT, 8, True, True)         '11100000
        Debug.Print BitShiftAriStrBin("10101011", 6, LEFT_SHIFT, 8, True, True)         '11000000
        Debug.Print BitShiftAriStrBin("10101011", 7, LEFT_SHIFT, 8, True, True)         '10000000
        Debug.Print BitShiftAriStrBin("10101011", 8, LEFT_SHIFT, 8, True, True)         '10000000
    '   Debug.Print "<<test 002-03>>"                                                   '<<test 002-03>>
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 8, True, True)         '�v���O������~
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 5, True, True)         '�v���O������~
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 4, True, True)         '�v���O������~
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 3, True, True)         '�v���O������~
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 2, True, True)         '�v���O������~
    '   Debug.Print BitShiftAriStrBin("0101011", 0, RIGHT_SHIFT, 1, True, True)         '�v���O������~
    '   Debug.Print BitShiftAriStrBin("10001011", 1, 5, 10, True, True)                 '�v���O������~
    '   Debug.Print BitShiftAriStrBin("", 2, LEFT_SHIFT, 10, True, True)                '�v���O������~
    '   Debug.Print BitShiftAriStrBin("10001011", -1, LEFT_SHIFT, 10, True, True)       '�v���O������~
    '   Debug.Print BitShiftAriStrBin("1000101A", 1, LEFT_SHIFT, 10, True, True)        '�v���O������~
    '   Debug.Print BitShiftAriStrBin("10001011", 1, LEFT_SHIFT, -1, True, True)        '�v���O������~
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = �T�v    �t�H���g�F/�w�i�F��ԋp����(�F�w���)
' = ����    rTrgtRange  Range     [in]  �Z��
' = ����    sColorType  String    [in]  �F��ʁiR or G or B�j
' = ����    sColorKind  Long      [in]  �F��ʁi1:Font�A2:Interior�j
' = ����    bIsHex      Boolean   [in]  ��i0:Decimal�A1:Hex�j
' = �ߒl                Variant         �w�i�F
' = �o��    �Ȃ�
' = �ˑ�    Funcs.bas/ConvRgb2X()
' = ����    Funcs.bas
' ==================================================================
Private Function GetColor( _
    ByRef rTrgtCell As Range, _
    ByVal sColorType As String, _
    ByVal lColorKind As Long, _
    ByVal bIsHex As Boolean _
) As Variant
    Dim lColorRGB As Long
    Dim lColorX As Long
    
    If rTrgtCell.Count > 1 Then
        GetColor = CVErr(xlErrRef)
    Else
        If lColorKind = 1 Then
            lColorRGB = rTrgtCell.Font.Color
        ElseIf lColorKind = 2 Then
            lColorRGB = rTrgtCell.Interior.Color
        Else
            GetColor = CVErr(xlErrValue)
        End If
        lColorX = ConvRgb2X(lColorRGB, sColorType)
        If lColorX > 255 Then
            GetColor = CVErr(xlErrValue)
        Else
            If bIsHex = True Then
                GetColor = UCase(String(2 - Len(Hex(lColorX)), "0") & Hex(lColorX))
            Else
                GetColor = lColorX
            End If
        End If
    End If
End Function
    Private Sub Test_GetColor()
        Dim oTrgtCellsPos01 As Range
        Dim oTrgtCellsNeg01 As Range
        Dim oTrgtCellsNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtCellsPos01 = .Cells(1, 1)
            Set oTrgtCellsNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
            Set oTrgtCellsNeg02 = .Range(.Cells(1, 1), .Cells(4, 1))
        End With
        
        Dim lColorIBefore As Long
        Dim lColorFBefore As Long
        lColorIBefore = oTrgtCellsPos01.Interior.Color
        lColorFBefore = oTrgtCellsPos01.Font.Color
        
        Debug.Print "*** test start! ***"
        Debug.Print "# Font Color"
        oTrgtCellsPos01.Font.Color = RGB(0, 0, 0)
        Debug.Print GetColor(oTrgtCellsPos01, "r", 1, False)  '0
        Debug.Print GetColor(oTrgtCellsPos01, "g", 1, False)  '0
        Debug.Print GetColor(oTrgtCellsPos01, "b", 1, False)  '0
        oTrgtCellsPos01.Font.Color = RGB(100, 100, 100)
        Debug.Print GetColor(oTrgtCellsPos01, "r", 1, False)  '100
        Debug.Print GetColor(oTrgtCellsPos01, "g", 1, False)  '100
        Debug.Print GetColor(oTrgtCellsPos01, "b", 1, False)  '100
        oTrgtCellsPos01.Font.Color = RGB(255, 255, 255)
        Debug.Print GetColor(oTrgtCellsPos01, "r", 1, False)  '255
        Debug.Print GetColor(oTrgtCellsPos01, "g", 1, False)  '255
        Debug.Print GetColor(oTrgtCellsPos01, "b", 1, False)  '255
        oTrgtCellsPos01.Font.Color = RGB(16, 100, 152)
        Debug.Print GetColor(oTrgtCellsPos01, "r", 1, False)  '16
        Debug.Print GetColor(oTrgtCellsPos01, "g", 1, False)  '100
        Debug.Print GetColor(oTrgtCellsPos01, "b", 1, False)  '152
        Debug.Print GetColor(oTrgtCellsPos01, "r", 1, True)   '10
        Debug.Print GetColor(oTrgtCellsPos01, "g", 1, True)   '64
        Debug.Print GetColor(oTrgtCellsPos01, "b", 1, True)   '98
        Debug.Print GetColor(oTrgtCellsPos01, "", 1, False)   '�G���[ 2015
        Debug.Print GetColor(oTrgtCellsPos01, "aa", 1, False) '�G���[ 2015
        Debug.Print GetColor(oTrgtCellsNeg01, "r", 1, False)  '�G���[ 2023
        Debug.Print GetColor(oTrgtCellsNeg02, "r", 1, False)  '�G���[ 2023
        
        Debug.Print "# Interior Color"
        oTrgtCellsPos01.Interior.Color = RGB(0, 0, 0)
        Debug.Print GetColor(oTrgtCellsPos01, "r", 2, False)  '0
        Debug.Print GetColor(oTrgtCellsPos01, "g", 2, False)  '0
        Debug.Print GetColor(oTrgtCellsPos01, "b", 2, False)  '0
        oTrgtCellsPos01.Interior.Color = RGB(100, 100, 100)
        Debug.Print GetColor(oTrgtCellsPos01, "r", 2, False)  '100
        Debug.Print GetColor(oTrgtCellsPos01, "g", 2, False)  '100
        Debug.Print GetColor(oTrgtCellsPos01, "b", 2, False)  '100
        oTrgtCellsPos01.Interior.Color = RGB(255, 255, 255)
        Debug.Print GetColor(oTrgtCellsPos01, "r", 2, False)  '255
        Debug.Print GetColor(oTrgtCellsPos01, "g", 2, False)  '255
        Debug.Print GetColor(oTrgtCellsPos01, "b", 2, False)  '255
        oTrgtCellsPos01.Interior.Color = RGB(16, 100, 152)
        Debug.Print GetColor(oTrgtCellsPos01, "r", 2, False)  '16
        Debug.Print GetColor(oTrgtCellsPos01, "g", 2, False)  '100
        Debug.Print GetColor(oTrgtCellsPos01, "b", 2, False)  '152
        Debug.Print GetColor(oTrgtCellsPos01, "r", 2, True)   '10
        Debug.Print GetColor(oTrgtCellsPos01, "g", 2, True)   '64
        Debug.Print GetColor(oTrgtCellsPos01, "b", 2, True)   '98
        Debug.Print GetColor(oTrgtCellsPos01, "", 2, False)   '�G���[ 2015
        Debug.Print GetColor(oTrgtCellsPos01, "aa", 2, False) '�G���[ 2015
        Debug.Print GetColor(oTrgtCellsNeg01, "r", 2, False)  '�G���[ 2023
        Debug.Print GetColor(oTrgtCellsNeg02, "r", 2, False)  '�G���[ 2023
        
        Debug.Print GetColor(oTrgtCellsNeg02, "r", 3, False)  '�G���[ 2023
        Debug.Print "*** test finished! ***"
        
        oTrgtCellsPos01.Interior.Color = lColorIBefore
        oTrgtCellsPos01.Interior.Color = lColorFBefore
    End Sub

' ==================================================================
' = �T�v    �t�H���g�F/�w�i�F��ԋp����(�S�F�w���)
' = ����    rTrgtRange  Range     [in]  �Z��
' = ����    lColorKind  Long      [in]  �F���(1:Font�A2:Interior)
' = ����    bIsHex      Boolean   [in]  ��i0:Decimal�A1:Hex�j
' = �ߒl                Variant         �w�i�F
' = �o��    �Ȃ�
' = �ˑ�    Funcs.bas/ConvRgb2X()
' = ����    Funcs.bas
' ==================================================================
Private Function GetColorAll( _
    ByRef rTrgtCell As Range, _
    ByVal lColorKind As Long, _
    ByVal bIsHex As Boolean _
) As Variant
    Dim lColorRGB As Long
    Dim lColorR As Long
    Dim lColorG As Long
    Dim lColorB As Long
    
    If rTrgtCell.Count > 1 Then
        GetColorAll = CVErr(xlErrRef)
    Else
        If lColorKind = 1 Then
            lColorRGB = rTrgtCell.Font.Color
        ElseIf lColorKind = 2 Then
            lColorRGB = rTrgtCell.Interior.Color
        Else
            GetColorAll = CVErr(xlErrValue)
        End If
        lColorR = ConvRgb2X(lColorRGB, "R")
        lColorG = ConvRgb2X(lColorRGB, "G")
        lColorB = ConvRgb2X(lColorRGB, "B")
        If lColorR > 255 Or lColorG > 255 Or lColorB > 255 Then
            GetColorAll = CVErr(xlErrValue)
        Else
            If bIsHex = True Then
                GetColorAll = UCase(String(2 - Len(Hex(lColorR)), "0") & Hex(lColorR)) & _
                        "," & UCase(String(2 - Len(Hex(lColorG)), "0") & Hex(lColorG)) & _
                        "," & UCase(String(2 - Len(Hex(lColorB)), "0") & Hex(lColorB))
            Else
                GetColorAll = lColorR & "," & lColorG & "," & lColorB
            End If
        End If
    End If
End Function
    Private Sub Test_GetColorAll()
        Dim oTrgtCellsPos01 As Range
        Dim oTrgtCellsNeg01 As Range
        Dim oTrgtCellsNeg02 As Range
        With ThisWorkbook.Sheets(1)
            Set oTrgtCellsPos01 = .Cells(1, 1)
            Set oTrgtCellsNeg01 = .Range(.Cells(1, 1), .Cells(1, 2))
            Set oTrgtCellsNeg02 = .Range(.Cells(1, 1), .Cells(4, 1))
        End With
        
        Dim lColorIBefore As Long
        Dim lColorFBefore As Long
        lColorIBefore = oTrgtCellsPos01.Interior.Color
        lColorFBefore = oTrgtCellsPos01.Font.Color
        
        Debug.Print "*** test start! ***"
        Debug.Print "# Font Color"
        oTrgtCellsPos01.Font.Color = RGB(0, 0, 0)
        Debug.Print GetColorAll(oTrgtCellsPos01, 1, False)   '0,0,0
        oTrgtCellsPos01.Font.Color = RGB(100, 100, 100)
        Debug.Print GetColorAll(oTrgtCellsPos01, 1, False)  '100,100,100
        oTrgtCellsPos01.Font.Color = RGB(255, 255, 255)
        Debug.Print GetColorAll(oTrgtCellsPos01, 1, False)  '255,255,255
        oTrgtCellsPos01.Font.Color = RGB(16, 100, 152)
        Debug.Print GetColorAll(oTrgtCellsPos01, 1, False)  '16,100,152
        Debug.Print GetColorAll(oTrgtCellsPos01, 1, True)   '10,64,98
        Debug.Print GetColorAll(oTrgtCellsNeg01, 1, False)  '�G���[ 2023
        Debug.Print GetColorAll(oTrgtCellsNeg02, 1, False)  '�G���[ 2023
        
        Debug.Print "# Interior Color"
        oTrgtCellsPos01.Interior.Color = RGB(0, 0, 0)
        Debug.Print GetColorAll(oTrgtCellsPos01, 2, False)  '0,0,0
        oTrgtCellsPos01.Interior.Color = RGB(100, 100, 100)
        Debug.Print GetColorAll(oTrgtCellsPos01, 2, False)  '100,100,100
        oTrgtCellsPos01.Interior.Color = RGB(255, 255, 255)
        Debug.Print GetColorAll(oTrgtCellsPos01, 2, False)  '255,255,255
        oTrgtCellsPos01.Interior.Color = RGB(16, 100, 152)
        Debug.Print GetColorAll(oTrgtCellsPos01, 2, False)  '16,100,152
        Debug.Print GetColorAll(oTrgtCellsPos01, 2, True)   '10,64,98
        Debug.Print GetColorAll(oTrgtCellsNeg01, 2, False)  '�G���[ 2023
        Debug.Print GetColorAll(oTrgtCellsNeg02, 2, False)  '�G���[ 2023
        
        Debug.Print GetColorAll(oTrgtCellsNeg02, 3, True)   '�G���[ 2023
        Debug.Print "*** test finished! ***"
        
        oTrgtCellsPos01.Interior.Color = lColorIBefore
        oTrgtCellsPos01.Interior.Color = lColorFBefore
    End Sub

' ==================================================================
' = �T�v    RGB�l�ϊ�
' = ����    lColorRGB       Long             [in]    RGB�l
' = ����    sColorType      String           [in]    �F���(r/g/b)
' = �ߒl                    Long                     �ϊ�����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Funcs.bas
' ==================================================================
Private Function ConvRgb2X( _
    ByVal lColorRGB As Long, _
    ByVal sColorType As String _
) As Long
    Select Case LCase(sColorType)
        Case "r": ConvRgb2X = lColorRGB Mod 256
        Case "g": ConvRgb2X = Int(lColorRGB / 256) Mod 256
        Case "b": ConvRgb2X = Int(lColorRGB / 256 / 256)
        Case Else: ConvRgb2X = 256 '�G���[
    End Select
End Function
    Private Sub Test_ConvRgb2X()
        Debug.Print "*** test start! ***"
        Debug.Print ConvRgb2X(RGB(255, 255, 255), "r") '255
        Debug.Print ConvRgb2X(RGB(255, 255, 255), "g") '255
        Debug.Print ConvRgb2X(RGB(255, 255, 255), "b") '255
        Debug.Print ConvRgb2X(RGB(0, 0, 0), "r")       '0
        Debug.Print ConvRgb2X(RGB(0, 0, 0), "g")       '0
        Debug.Print ConvRgb2X(RGB(0, 0, 0), "b")       '0
        Debug.Print ConvRgb2X(RGB(16, 39, 40), "r")    '16
        Debug.Print ConvRgb2X(RGB(16, 39, 40), "g")    '39
        Debug.Print ConvRgb2X(RGB(16, 39, 40), "b")    '40
        Debug.Print ConvRgb2X(RGB(16, 39, 40), "")     '256
        Debug.Print ConvRgb2X(RGB(16, 39, 40), "a")    '256
        Debug.Print ConvRgb2X(RGB(16, 39, 40), "aa")   '256
        Debug.Print "*** test finished! ***"
    End Sub

