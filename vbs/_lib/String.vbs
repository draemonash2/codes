Option Explicit

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕������ԋp����B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ���o������
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    String.vbs
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr, _
    ByVal sDlmtr _
)
    Dim asSplitWord
    
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function
    'Call Test_ExtractTailWord()
    Private Sub Test_ExtractTailWord()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "\" )   ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a", "\" )       ' a
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\", "\" )        ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\" )         ' test
        Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\\" )        ' C:\test
        Result = Result & vbNewLine & ExtractTailWord( "a.txt", "\" )           ' a.txt
        Result = Result & vbNewLine & ExtractTailWord( "", "\" )                ' 
        Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "" )    ' C:\test\a.txt
        Result = Result & vbNewLine & "*** test finished! ***"
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕��������������B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ����������
' = �o��    �Ȃ�
' = �ˑ�    String.vbs/ExtractTailWord()
' = ����    String.vbs
' ==================================================================
Public Function RemoveTailWord( _
    ByVal sStr, _
    ByVal sDlmtr _
)
    Dim sTailWord
    Dim lRemoveLen
    
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
                RemoveTailWord = Left(sStr, Len(sStr) - lRemoveLen)
            End If
        End If
    End If
End Function
    'Call Test_RemoveTailWord()
    Private Sub Test_RemoveTailWord()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & "*** test start! ***"
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "\" )    ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a", "\" )        ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\", "\" )         ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\" )          ' C:
        Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\\" )         ' C:\test
        Result = Result & vbNewLine & RemoveTailWord( "", "\" )                 ' 
        Result = Result & vbNewLine & RemoveTailWord( "a.txt", "\" )            ' a.txt�i�t�@�C�������ǂ����͔��f���Ȃ��j
        Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "" )     ' C:\test\a.txt
        Result = Result & vbNewLine & "*** test finished! ***"
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�H���_�p�X
' = �o��    ���[�J���t�@�C���p�X�i��Fc:\test�j�� URL �i��Fhttps://test�j
' =         ���w��\
' = �ˑ�    String.vbs/RemoveTailWord()
' = ����    String.vbs
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath _
)
    If InStr( sFilePath, "\" ) Then
        GetDirPath = RemoveTailWord( sFilePath, "\" )
    ElseIf InStr( sFilePath, "/" ) Then
        GetDirPath = RemoveTailWord( sFilePath, "/" )
    Else
        GetDirPath = sFilePath
    End If
End Function
    'Call Test_GetDirPath()
    Private Sub Test_GetDirPath()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetDirPath( "C:\test\a.txt" )    ' C:\test
        Result = Result & vbNewLine & GetDirPath( "http://test/a" )    ' http://test
        Result = Result & vbNewLine & GetDirPath( "C:_test_a.txt" )    ' C:_test_a.txt
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�@�C�����𒊏o����
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�@�C����
' = �o��    ���[�J���t�@�C���p�X�i��Fc:\test�j�� URL �i��Fhttps://test�j
' =         ���w��\
' = �ˑ�    String.vbs/ExtractTailWord()
' = ����    String.vbs
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath _
)
    If InStr( sFilePath, "\" ) Then
        GetFileName = ExtractTailWord( sFilePath, "\" )
    ElseIf InStr( sFilePath, "/" ) Then
        GetFileName = ExtractTailWord( sFilePath, "/" )
    Else
        GetFileName = sFilePath
    End If
End Function
    'Call Test_GetFileName()
    Private Sub Test_GetFileName()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileName( "C:\test\a.txt" )    ' a.txt
        Result = Result & vbNewLine & GetFileName( "http://test/a" )    ' a
        Result = Result & vbNewLine & GetFileName( "c:_test_a" )        ' c:_test_a
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����g���q�𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �g���q
' = �o��    �E�g���q���Ȃ��ꍇ�A�󕶎���ԋp����
' =         �E�t�@�C�������w��\
' = �ˑ�    String.vbs/ExtractTailWord()
' =         String.vbs/GetFileName()
' = ����    String.vbs
' ==================================================================
Public Function GetFileExt( _
    ByVal sFilePath _
)
    Dim sFileName
    sFileName = GetFileName(sFilePath)
    If InStr(sFileName, ".") > 0 Then
        GetFileExt = ExtractTailWord(sFileName, ".")
    Else
        GetFileExt = ""
    End If
End Function
    'Call Test_GetFileExt()
    Private Sub Test_GetFileExt()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileExt("c:\codes\test.txt")     'txt
        Result = Result & vbNewLine & GetFileExt("c:\codes\test")         '
        Result = Result & vbNewLine & GetFileExt("test.txt")              'txt
        Result = Result & vbNewLine & GetFileExt("test")                  '
        Result = Result & vbNewLine & GetFileExt("c:\codes\test.aaa.txt") 'txt
        Result = Result & vbNewLine & GetFileExt("test.aaa.txt")          'txt
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�@�C���x�[�X���𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�@�C���x�[�X��
' = �o��    �E�g���q���Ȃ��ꍇ�A�󕶎���ԋp����
' =         �E�t�@�C�������w��\
' = �ˑ�    String.vbs/RemoveTailWord()
' =         String.vbs/GetFileName()
' = ����    String.vbs
' ==================================================================
Public Function GetFileBase( _
    ByVal sFilePath _
)
    Dim sFileName
    sFileName = GetFileName(sFilePath)
    GetFileBase = RemoveTailWord(sFileName, ".")
End Function
    'Call Test_GetFileBase()
    Private Sub Test_GetFileBase()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFileBase("c:\codes\test.txt")     'test
        Result = Result & vbNewLine & GetFileBase("c:\codes\test")         'test
        Result = Result & vbNewLine & GetFileBase("test.txt")              'test
        Result = Result & vbNewLine & GetFileBase("test")                  'test
        Result = Result & vbNewLine & GetFileBase("c:\codes\test.aaa.txt") 'test.aaa
        Result = Result & vbNewLine & GetFileBase("test.aaa.txt")          'test.aaa
        MsgBox Result
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
' = �ˑ�    String.vbs/GetDirPath()
' =         String.vbs/GetFileName()
' =         String.vbs/GetFileBase()
' =         String.vbs/GetFileExt()
' = ����    String.vbs
' ==================================================================
Public Function GetFilePart( _
    ByVal sFilePath, _
    ByVal lPartType _
)
    Select Case lPartType
        Case 1: GetFilePart = GetDirPath(sFilePath)
        Case 2: GetFilePart = GetFileName(sFilePath)
        Case 3: GetFilePart = GetFileBase(sFilePath)
        Case 4: GetFilePart = GetFileExt(sFilePath)
        Case Else: GetFilePart = ""
    End Select
End Function
    'Call Test_GetFilePart()
    Private Sub Test_GetFilePart()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 0)     '
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 1)     'c:\codes
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 2)     'test.txt
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 3)     'test
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 4)     'txt
        Result = Result & vbNewLine & GetFilePart("c:\codes\test.txt", 5)     '
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �����`����ϊ�����B�i��F2017/03/22 18:20:14 �� 20170322-182014�j
' = ����    sDateBfr        String  [in]    �ϊ��O ����������
' = ����    lNamingRuleType Long    [in]    �����K�����
' =                                           1(�Z�k���[�h)  : YYMMDD-HHMM
' =                                           2(�Z�k���[�h2) : YYMMDD
' =                                           2(�Z�k���[�h3) : YYMMDDHHMM
' =                                           other          : YYYYMMDD-HHMMSS
' = �ߒl                    String          �ϊ��� ����������
' = �o��    �E��ɓ������t�@�C������t�H���_���Ɏg�p����ۂɎg�p����B
' =         �E�����̌`��������Ă����ꍇ(ex.2017/13/22 18:25:14)�A�󕶎����Ԃ��B
' = �ˑ�    �Ȃ�
' = ����    String.vbs
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateBfr, _
    ByVal lNamingRuleType _
)
    On Error Resume Next
    Dim sDateAfr
    Select Case lNamingRuleType
        Case 1  '�Z�k���[�h
            sDateAfr = _
                Right(Year(sDateBfr), 2) & _
                String(2 - Len(Month(sDateBfr)),  "0") & Month(sDateBfr)  & _
                String(2 - Len(Day(sDateBfr)),    "0") & Day(sDateBfr)    & _
                "-" & _
                String(2 - Len(Hour(sDateBfr)),   "0") & Hour(sDateBfr)   & _
                String(2 - Len(Minute(sDateBfr)), "0") & Minute(sDateBfr)
        Case 2  '�Z�k���[�h2
            sDateAfr = _
                Right(Year(sDateBfr), 2) & _
                String(2 - Len(Month(sDateBfr)),  "0") & Month(sDateBfr)  & _
                String(2 - Len(Day(sDateBfr)),    "0") & Day(sDateBfr)
        Case 3  '�Z�k���[�h3
            sDateAfr = _
                Right(Year(sDateBfr), 2) & _
                String(2 - Len(Month(sDateBfr)),  "0") & Month(sDateBfr)  & _
                String(2 - Len(Day(sDateBfr)),    "0") & Day(sDateBfr)    & _
                String(2 - Len(Hour(sDateBfr)),   "0") & Hour(sDateBfr)   & _
                String(2 - Len(Minute(sDateBfr)), "0") & Minute(sDateBfr)
        Case Else
            sDateAfr = _
                String(4 - Len(Year(sDateBfr)),   "0") & Year(sDateBfr)   & _
                String(2 - Len(Month(sDateBfr)),  "0") & Month(sDateBfr)  & _
                String(2 - Len(Day(sDateBfr)),    "0") & Day(sDateBfr)    & _
                "-" & _
                String(2 - Len(Hour(sDateBfr)),   "0") & Hour(sDateBfr)   & _
                String(2 - Len(Minute(sDateBfr)), "0") & Minute(sDateBfr) & _
                String(2 - Len(Second(sDateBfr)), "0") & Second(sDateBfr)
    End Select
    If Err.Number = 0 Then
        ConvDate2String = sDateAfr
    Else
        ConvDate2String = ""
    End If
    On Error Goto 0
End Function
    'Call Test_ConvDate2String()
    Private Sub Test_ConvDate2String()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & ConvDate2String("2001/12/31 21:23:45", 0) ' 20011231-212345
        Result = Result & vbNewLine & ConvDate2String("2001/1/1 21:23:45", 0)   ' 20010101-212345
        Result = Result & vbNewLine & ConvDate2String("1986/01/01 0:0:0", 0)    ' 19860101-000000
        Result = Result & vbNewLine & ConvDate2String("2001/12/31", 0)          ' 20011231-000000
        Result = Result & vbNewLine & ConvDate2String("2001/12/31 21:23:45", 1) ' 011231-2123
        Result = Result & vbNewLine & ConvDate2String("2001/1/1 21:23:45", 1)   ' 010101-2123
        Result = Result & vbNewLine & ConvDate2String("1986/01/01 0:0:0", 1)    ' 860101-0000
        Result = Result & vbNewLine & ConvDate2String("2001/12/31", 1)          ' 011231-0000
        Result = Result & vbNewLine & ConvDate2String("2001/12/31 21:23:45", 2) ' 011231
        Result = Result & vbNewLine & ConvDate2String("2001/1/1 21:23:45", 2)   ' 010101
        Result = Result & vbNewLine & ConvDate2String("1986/01/01 0:0:0", 2)    ' 860101
        Result = Result & vbNewLine & ConvDate2String("2001/12/31", 2)          ' 011231
        Result = Result & vbNewLine & ConvDate2String("2001/12/31 21:23:45", 3) ' 0112312123
        Result = Result & vbNewLine & ConvDate2String("2001/1/1 21:23:45", 3)   ' 0101012123
        Result = Result & vbNewLine & ConvDate2String("1986/01/01 0:0:0", 3)    ' 8601010000
        Result = Result & vbNewLine & ConvDate2String("2001/12/31", 3)          ' 0112310000
        Result = Result & vbNewLine & ConvDate2String("01/12/31 21:23:45", 0)   ' 20011231-212345
        Result = Result & vbNewLine & ConvDate2String("96/12/31 21:23:45", 0)   ' 19960101-212345
        
        Result = Result & vbNewLine & ConvDate2String("2001/0/31 21:23:45", 0)   ' 
        Result = Result & vbNewLine & ConvDate2String("2001/13/31 21:23:45", 0)  ' 
        Result = Result & vbNewLine & ConvDate2String("2001/12/32 21:23:45", 0)  ' 
        Result = Result & vbNewLine & ConvDate2String("2001/12/31 25:23:45", 0)  ' 
        Result = Result & vbNewLine & ConvDate2String("2001/12/31 21:60:45", 0)  ' 
        Result = Result & vbNewLine & ConvDate2String("2001/12/31 21:23:60", 0)  ' 
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �w�肳�ꂽ������̕����񒷁i�o�C�g���j��ԋp����
' = ����    sInStr      String  [in]  ������
' = �ߒl                Long          �����񒷁i�o�C�g���j
' = �o��    �W���ŗp�ӂ���Ă��� LenB() �֐��́AUnicode �ɂ�����
' =         �o�C�g����ԋp���邽�ߔ��p�������Q�����Ƃ��ăJ�E���g����B
' =           �i��FLenB("�t�@�C���T�C�Y ") �� 16�j
' =         ���̂��߁A���p�������P�����Ƃ��ăJ�E���g����{�֐���p�ӁB
' = �ˑ�    �Ȃ�
' = ����    String.vbs
' ==================================================================
Public Function LenByte( _
    ByVal sInStr _
)
    Dim lIdx, sChar
    LenByte = 0
    If Trim(sInStr) <> "" Then
        For lIdx = 1 To Len(sInStr)
            sChar = Mid(sInStr, lIdx, 1)
            '�Q�o�C�g�����́{�Q
            If (Asc(sChar) And &HFF00) <> 0 Then
                LenByte = LenByte + 2
            Else
                LenByte = LenByte + 1
            End If
        Next
    End If
End Function
    'Call Test_LenByte()
    Private Sub Test_LenByte()
        Dim Result
        Result = "[Result]"
        Result = Result & vbNewLine & LenByte( "aaa" )      ' 3
        Result = Result & vbNewLine & LenByte( "aaa " )     ' 4
        Result = Result & vbNewLine & LenByte( "" )         ' 0
        Result = Result & vbNewLine & LenByte( "������" )   ' 6
        Result = Result & vbNewLine & LenByte( "������ " )  ' 7
        Result = Result & vbNewLine & LenByte( "���� ��" )  ' 7
        Result = Result & vbNewLine & LenByte( Chr(9) )     ' 1
        Result = Result & vbNewLine & LenByte( Chr(10) )    ' 1
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �����񒷂���^�u����؂�ʒu�܂ł̃^�u��������ԋp����
' = ����    lLen            Long    [in]  ������
' = ����    lLenMax         Long    [in]  ������ő咷
' = ����    lTabWidth       Long    [in]  �^�u������
' = ����    lPaddingTabNum  Long    [out] �^�u������
' = �ߒl                    Boolean       ����(OK,NG)
' = �o��    ���s��)lLen:3,lLenMax:9,lTabWidth:4
'             �uxxx^^   ^   �v
'             �uxxxxxxxxx^  �v
'               ��return:3
' = �ˑ�    String.vbs/CalcPaddingWidth()
' = ����    String.vbs
' ==================================================================
Public Function CalcPaddingTabWidth( _
    ByVal lLen, _
    ByVal lLenMax, _
    ByVal lTabWidth, _
    ByRef lPaddingTabNum _
)
    Dim lPaddingLen
    Dim dCalcResult
    Dim lDigit
    If lTabWidth = 0 Then
        lPaddingTabNum = 0
        CalcPaddingTabWidth = False
    ElseIf lTabWidth < 0 Or lLen < 0 Or lLenMax < 0 Then
        lPaddingTabNum = 0
        CalcPaddingTabWidth = False
    Else
        '�p�f�B���O��(�X�y�[�X)�Z�o
        If CalcPaddingWidth(lLen, lLenMax, lTabWidth, lPaddingLen) Then
            '�p�f�B���O��(�^�u)�Z�o
            lDigit = 0
            dCalcResult = lPaddingLen / lTabWidth
            lPaddingTabNum = Fix((dCalcResult + (9 * (10 ^ (-1 * (lDigit + 1))))) * (10 ^ lDigit)) / (10 ^ lDigit)
            CalcPaddingTabWidth = True
        Else
            lPaddingTabNum = 0
            CalcPaddingTabWidth = False
        End If
    End If
End Function
    'Call Test_CalcPaddingTabWidth()
    Private Sub Test_CalcPaddingTabWidth()
        Dim Result
        Dim lPaddingLen
        Result = "[Result]"
        Result = Result & vbNewLine & CalcPaddingTabWidth(2, 0, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(2, 1, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(2, 4, 4,  lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & CalcPaddingTabWidth(2, 6, 4,  lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & CalcPaddingTabWidth(4, 0, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 0, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 2, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 3, 4,  lPaddingLen)  & ":" & lPaddingLen '1
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 4, 4,  lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & CalcPaddingTabWidth(0, 5, 4,  lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 19, 4, lPaddingLen)  & ":" & lPaddingLen '4
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 20, 4, lPaddingLen)  & ":" & lPaddingLen '5
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 21, 4, lPaddingLen)  & ":" & lPaddingLen '5
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 22, 4, lPaddingLen)  & ":" & lPaddingLen '5
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 23, 4, lPaddingLen)  & ":" & lPaddingLen '5
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 24, 4, lPaddingLen)  & ":" & lPaddingLen '6
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 15, 8, lPaddingLen)  & ":" & lPaddingLen '2
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 16, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 17, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 18, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 19, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 20, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 21, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 22, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 23, 8, lPaddingLen)  & ":" & lPaddingLen '3
        Result = Result & vbNewLine & CalcPaddingTabWidth(5, 24, 8, lPaddingLen)  & ":" & lPaddingLen '4
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingTabWidth(1, 5, 0,  lPaddingLen)  & ":" & lPaddingLen 'False
        Result = Result & vbNewLine & CalcPaddingTabWidth(1, 5, -1, lPaddingLen)  & ":" & lPaddingLen 'False
        Result = Result & vbNewLine & CalcPaddingTabWidth(1, -1, 4, lPaddingLen)  & ":" & lPaddingLen 'False
        Result = Result & vbNewLine & CalcPaddingTabWidth(-1, 5, 4, lPaddingLen)  & ":" & lPaddingLen 'False
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �����񒷂����؂蕝�ʒu�܂ł̕�������ԋp����
' = ����    lLen        Long    [in]  ������
' = ����    lLenMax     Long    [in]  ������ő咷
' = ����    lSepWidth   Long    [in]  ��؂蕝
' = ����    lPaddingLen Long    [out] ������
' = �ߒl                Boolean       ����(OK,NG)
' = �o��    ���s��)lLen:3,lLenMax:9,lSepWidth:4
'             �uxxx         �v
'             �uxxxxxxxxx   �v
'               ��return:9
' = �ˑ�    �Ȃ�
' = ����    String.vbs
' ==================================================================
Public Function CalcPaddingWidth( _
    ByVal lLen, _
    ByVal lLenMax, _
    ByVal lSepWidth, _
    ByRef lPaddingLen _
)
    Dim lPaddingWidth
    If lSepWidth = 0 Then
        lPaddingLen = 0
        CalcPaddingWidth = False
    ElseIf lSepWidth < 0 Or lLen < 0 Or lLenMax < 0 Then
        lPaddingLen = 0
        CalcPaddingWidth = False
    Else
        If lLen > lLenMax Then
            lLenMax = lLen
        End If
        lPaddingWidth = lSepWidth - (lLenMax Mod lSepWidth)
        lPaddingLen = (lPaddingWidth + lLenMax) - lLen
        CalcPaddingWidth = True
    End If
End Function
    'Call Test_CalcPaddingWidth()
    Private Sub Test_CalcPaddingWidth()
        Dim Result
        Dim lPaddingLen
        Result = "[Result]"
        Result = Result & vbNewLine & CalcPaddingWidth(0, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 8
        Result = Result & vbNewLine & CalcPaddingWidth(3, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 5
        Result = Result & vbNewLine & CalcPaddingWidth(4, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & CalcPaddingWidth(5, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 3
        Result = Result & vbNewLine & CalcPaddingWidth(6, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 2
        Result = Result & vbNewLine & CalcPaddingWidth(7, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 1
        Result = Result & vbNewLine & CalcPaddingWidth(8, 5, 4, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingWidth(0, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 8
        Result = Result & vbNewLine & CalcPaddingWidth(1, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(2, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 6
        Result = Result & vbNewLine & CalcPaddingWidth(3, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 5
        Result = Result & vbNewLine & CalcPaddingWidth(4, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & CalcPaddingWidth(5, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 3
        Result = Result & vbNewLine & CalcPaddingWidth(6, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 2
        Result = Result & vbNewLine & CalcPaddingWidth(7, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 1
        Result = Result & vbNewLine & CalcPaddingWidth(8, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 8
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingWidth(0, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(3, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & CalcPaddingWidth(4, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 3
        Result = Result & vbNewLine & CalcPaddingWidth(5, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 2
        Result = Result & vbNewLine & CalcPaddingWidth(6, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 1
        Result = Result & vbNewLine & CalcPaddingWidth(7, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(8, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 6
        Result = Result & vbNewLine & ""
        Result = Result & vbNewLine & CalcPaddingWidth(1, 5, 0, lPaddingLen)  & ":" & lPaddingLen ' False
        Result = Result & vbNewLine & CalcPaddingWidth(1, 5, -1, lPaddingLen) & ":" & lPaddingLen ' False
        Result = Result & vbNewLine & CalcPaddingWidth(1, -1, 4, lPaddingLen) & ":" & lPaddingLen ' False
        Result = Result & vbNewLine & CalcPaddingWidth(-1, 5, 4, lPaddingLen) & ":" & lPaddingLen ' False
        MsgBox Result
    End Sub

' ==================================================================
' = �T�v    �����L�[����l�𒊏o����B
' = ����    sTrgtStr    String      [in]  �Ώە�����
' = ����    asKeyList   String()    [in]  �L�[���X�g
' = ����    asValueList String()    [out] �l���X�g
' = ����    bTrimValue  Boolean     [in]  �擪�����󔒏���
' = �ߒl                Boolean           ��v�L��
' = �o��    �EasKeyList�͕K��1�I���W���Ŏw�肷�邱��
' = �o��    �E���s��͈ȉ��B
' =           ex1�j
' =             <<����>>
' =               sTrgtStr = "DESC:aBC FORM:d AL LSB:12 ALLOC:aaa END: "
' =               asKeyList(1) = "FORM:"
' =               asKeyList(2) = "LSB:"
' =               asKeyList(3) = "DESC:"
' =               asKeyList(4) = "END:"
' =               bTrimValue = True
' =             <<�o��>>
' =               asValueList(0) = ""
' =               asValueList(1) = "d AL"
' =               asValueList(2) = "12"
' =               asValueList(3) = "aBC"
' =               asValueList(4) = " "
' =           ex2�j
' =             <<����>>
' =               sTrgtStr = "DESC:aBC FORM:d AL LSB:12 ALLOC:aaa END: "
' =               asKeyList(1) = "FORM:"
' =               asKeyList(2) = "ALLOC:"
' =               bTrimValue = False
' =             <<�o��>>
' =               asValueList(0) = "aBC "
' =               asValueList(1) = "d AL LSB:12 "
' =               asValueList(2) = "aaa END: "
' = �ˑ�    �Ȃ�
' = ����    String.vbs
' ==================================================================
Public Function ExtractValuesFrKeys( _
    ByVal sTrgtStr, _
    ByRef asKeyList(), _
    ByRef asValueList(), _
    ByVal bTrimValue _
)
    Dim lKeyNum
    lKeyNum = UBound(asKeyList)
    
    ReDim asValueList(lKeyNum) '0�I���W��(0�̓q�b�g���Ȃ�����������)
    
    Dim alKeyMatchCharIdx()
    ReDim alKeyMatchCharIdx(lKeyNum)
    Dim lKeyIdx
    For lKeyIdx = 1 To lKeyNum
        alKeyMatchCharIdx(lKeyIdx) = 1
    Next
    
    Dim lCnfrmKeyIdxOld '0=���m��
    lCnfrmKeyIdxOld = 0
    
    Dim lCnfrmKeyStrPosOld
    lCnfrmKeyStrPosOld = 1
    
    ExtractValuesFrKeys = False
    
    Dim lCurStrPos
    For lCurStrPos = 1 To Len(sTrgtStr)
        Dim sCurChar
        sCurChar = Mid(sTrgtStr, lCurStrPos, 1)
        
        Dim lCnfrmKeyIdxNow '0=���m��
        lCnfrmKeyIdxNow = 0
        
        '*** �m�蔻�� ***
        For lKeyIdx = 1 To lKeyNum
            If asKeyList(lKeyIdx) = "" Then
                'Do Nothing
            Else
                Dim sCurKeyChar
                sCurKeyChar = Mid( _
                                    asKeyList(lKeyIdx), _
                                    alKeyMatchCharIdx(lKeyIdx), _
                                    1 _
                                )
                If sCurKeyChar = sCurChar Then
                    If Len(asKeyList(lKeyIdx)) <= alKeyMatchCharIdx(lKeyIdx) Then
                        lCnfrmKeyIdxNow = lKeyIdx
                        ExtractValuesFrKeys = True
                    Else
                        alKeyMatchCharIdx(lKeyIdx) = alKeyMatchCharIdx(lKeyIdx) + 1
                    End If
                Else
                    alKeyMatchCharIdx(lKeyIdx) = 1
                End If
                If lCnfrmKeyIdxNow > 0 Then
                    Exit For
                End If
            End If
        Next
        
        '*** �m�蔻��㎖�㏈�� ***
        Dim lExtractLen
        If lCnfrmKeyIdxNow > 0 Then
            If lCnfrmKeyIdxOld > 0 Then '�q�b�g���ڈȍ~
                lExtractLen = lCurStrPos - Len(asKeyList(lCnfrmKeyIdxNow)) - lCnfrmKeyStrPosOld + 1
                asValueList(lCnfrmKeyIdxOld) = _
                    Mid( _
                        sTrgtStr, _
                        lCnfrmKeyStrPosOld, _
                        lExtractLen _
                    )
                If bTrimValue = True Then
                    asValueList(lCnfrmKeyIdxOld) = Trim(asValueList(lCnfrmKeyIdxOld))
                End If
            Else '�q�b�g����
                If lCurStrPos > Len(asKeyList(lCnfrmKeyIdxNow)) Then
                    lExtractLen = lCurStrPos - Len(asKeyList(lCnfrmKeyIdxNow))
                    asValueList(0) = _
                        Mid( _
                            sTrgtStr, _
                            1, _
                            lExtractLen _
                        )
                    If bTrimValue = True Then
                        asValueList(0) = Trim(asValueList(0))
                    End If
                End If
            End If
            
            '�N���A
            For lKeyIdx = 1 To lKeyNum
                alKeyMatchCharIdx(lKeyIdx) = 1
            Next
            
            '�O��l�X�V
            lCnfrmKeyIdxOld = lCnfrmKeyIdxNow
            lCnfrmKeyStrPosOld = lCurStrPos + 1
        End If
    Next
    
    '�ŏI�v�f���o��
    If lCnfrmKeyIdxOld > 0 Then '1��ȏ�q�b�g
        lExtractLen = Len(sTrgtStr) - lCnfrmKeyStrPosOld + 1
        asValueList(lCnfrmKeyIdxOld) = _
            Mid( _
                sTrgtStr, _
                lCnfrmKeyStrPosOld, _
                lExtractLen _
            )
        If bTrimValue = True Then
            asValueList(lCnfrmKeyIdxOld) = Trim(asValueList(lCnfrmKeyIdxOld))
        End If
    Else '1����q�b�g����
        asValueList(0) = sTrgtStr
        If bTrimValue = True Then
            asValueList(0) = Trim(asValueList(0))
        End If
    End If
End Function
    'Call Test_ExtractValuesFrKeys()
    Private Sub Test_ExtractValuesFrKeys()
        Dim asKeyList()
        Dim sOutMsg
        
        '����n
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        asKeyList(2) = "LSB:"
        asKeyList(3) = "FORM:"
        asKeyList(4) = "MONI:"
        asKeyList(5) = "ALLOC:"
        asKeyList(6) = "END:"
        Call TestSub_ExtractValuesFrKeysTrim("DESC:aBC FORM:d AL LSB:12 ALLOC:aaa END:", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("DESC: aBC FORM:d AL LsB:12 ALLOC:aaa END:", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("DESC:aBC FORM:d AL LSB;12 ALLOC:aaa END:", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END: ", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("d AL DESC:aBC ", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("d AL DeSC:aBC ", asKeyList, False, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("DESC:aBC FORM:d AL LSB:12 DESC:DEf ALLOC:aaa END:", asKeyList, False, sOutMsg)
        
        Call TestSub_ExtractValuesFrKeysTrim("DESC: aBC FORM:d AL LsB:12 ALLOC:aaa END:", asKeyList, True, sOutMsg)
        Call TestSub_ExtractValuesFrKeysTrim("DESC: aBC FORM:d AL LsB:12 ALLOC:aaa END:", asKeyList, False, sOutMsg)
        
        '�ُ�n(asKeyList��������)
        'Dim asKeyList2()
        'Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList2, False, sOutMsg)
        
        '�ُ�n(asKeyList�ɋ󔒗v�f����)
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        asKeyList(3) = "LSB:"
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        ReDim asKeyList(6)
        asKeyList(3) = "LSB:"
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        ReDim asKeyList(6)
        asKeyList(3) = "END:"
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        
        '�ُ�n(asKeyList���󔒗v�f�̂�)
        ReDim asKeyList(6)
        Call TestSub_ExtractValuesFrKeysTrim("FORM:d AL LSB:12 ALLOC:bbb DESC:aBC END:ab", asKeyList, False, sOutMsg)
        
        '�ُ�n(sTrgtStr�󕶎���)
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        asKeyList(2) = "LSB:"
        asKeyList(3) = "FORM:"
        asKeyList(4) = "MONI:"
        asKeyList(5) = "ALLOC:"
        asKeyList(6) = "END:"
        Call TestSub_ExtractValuesFrKeysTrim("", asKeyList, False, sOutMsg)
        
        '�ُ�n(�����񒷕s��)
        ReDim asKeyList(6)
        asKeyList(1) = "DESC:"
        asKeyList(2) = "LSB:"
        asKeyList(3) = "FORM:"
        asKeyList(4) = "MONI:"
        asKeyList(5) = "ALLOC:"
        asKeyList(6) = "END:"
        Call TestSub_ExtractValuesFrKeysTrim("a", asKeyList, False, sOutMsg)
        Msgbox sOutMsg
    End Sub
    Private Function TestSub_ExtractValuesFrKeysTrim( _
        ByVal sTrgtStr, _
        ByRef asKeyList(), _
        ByVal bTrimValue, _
        ByRef sOutMsg _
    )
        Dim asValueList()
        Dim bRet
        Redim asValueList(0)
        bRet = ExtractValuesFrKeys(sTrgtStr, asKeyList, asValueList, bTrimValue)
        sOutMsg = sOutMsg & vbNewLine & sTrgtStr
        sOutMsg = sOutMsg & vbNewLine & bRet
        
        Dim lIdx
        For lIdx = 0 To UBound(asValueList)
            If lIdx <= 0 Then
                sOutMsg = sOutMsg & vbNewLine & "other:""" & asValueList(lIdx) & """"
            Else
                sOutMsg = sOutMsg & vbNewLine & asKeyList(lIdx) & """" & asValueList(lIdx) & """"
            End If
        Next
        sOutMsg = sOutMsg & vbNewLine & ""
    End Function

' ==================================================================
' = �T�v    ��΃p�X���猟���L�[�z���K�w�̑��΃p�X�֒u��
' = ����    sInFilePath     String  [in]    ��΃p�X
' = ����    sMatchDirName   String  [in]    �����Ώۃt�H���_��
' = ����    lRemeveDirLevel Long    [in]    �K�w���x��
' = ����    sRelativePath   String  [out]   ���΃p�X
' = �ߒl                    Boolean         ��������
' = �o��    ���s��1)
'             sInFilePath     : c\codes\aaa\bbb\ccc\test.txt
'             sMatchDirName   : codes
'             lRemeveDirLevel : 1
'             ��
'             sRelativePath   : bbb\ccc\test.txt
'             �ߒl            : true
'
'           ���s��2)
'             sInFilePath     : c\codes\aaa\bbb\ccc\test.txt
'             sMatchDirName   : code
'             lRemeveDirLevel : 2
'             ��
'             sRelativePath   : c\codes\aaa\bbb\ccc\test.txt
'             �ߒl            : false
' = �ˑ�    �Ȃ�
' = ����    String.vbs
' ==================================================================
Public Function ExtractRelativePath( _
    ByVal sInFilePath, _
    ByVal sMatchDirName, _
    ByVal lRemeveDirLevel, _
    ByRef sRelativePath _
)
    Dim sRemoveDirLevelPath
    sRemoveDirLevelPath = ""
    Dim lIdx
    For lIdx = 0 To lRemeveDirLevel - 1
        sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
    Next
    
    Dim sSearchPattern
    Dim sTargetStr
    sSearchPattern = ".*\\" & sMatchDirName & sRemoveDirLevelPath & "\\"
    sTargetStr = sInFilePath
    
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSearchPattern                '�����p�^�[����ݒ�
    oRegExp.IgnoreCase = True                       '�啶���Ə���������ʂ��Ȃ�
    oRegExp.Global = True                           '������S�̂�����
    
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)  '�p�^�[���}�b�`���s
    
    If oMatchResult.Count > 0 Then
        sRelativePath = Replace(sInFilePath, oMatchResult.Item(0), "")
        ExtractRelativePath = True
    Else
        sRelativePath = sInFilePath
        ExtractRelativePath = False
    End If
End Function
    'Call Test_ExtractRelativePath()
    Private Sub Test_ExtractRelativePath()
        Dim sResult
        Dim sInFilePath
        Dim sOutFilePath
        sInFilePath = "c:\test\a\aa\bbb\test.txt"
        sOutFilePath = ""
        sResult = "[Result]"
        sResult = sResult & vbNewLine & ExtractRelativePath( sInFilePath, "a", 1, sOutFilePath )          & "�F" & sOutFilePath    ' True�Fbbb\test.txt
        sResult = sResult & vbNewLine & ExtractRelativePath( sInFilePath, "a", 0, sOutFilePath )          & "�F" & sOutFilePath    ' True�Faa\bbb\test.txt
        sResult = sResult & vbNewLine & ExtractRelativePath( sInFilePath, "a", -1, sOutFilePath )         & "�F" & sOutFilePath    ' True�Faa\bbb\test.txt
        sResult = sResult & vbNewLine & ExtractRelativePath( sInFilePath, "a", 5, sOutFilePath )          & "�F" & sOutFilePath    ' False�Fc:\test\a\aa\bbb\test.txt
        sResult = sResult & vbNewLine & ExtractRelativePath( sInFilePath, "test.txt", 0, sOutFilePath )   & "�F" & sOutFilePath    ' False�Fc:\test\a\aa\bbb\test.txt
        MsgBox sResult
    End Sub
