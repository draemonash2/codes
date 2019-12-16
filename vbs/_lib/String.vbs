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
' = ����    sDateTime   String  [in]  �����iYYYY/MM/DD HH:MM:SS�j
' = �ߒl                String        �����iYYYYMMDD-HHMMSS�j
' = �o��    ��ɓ������t�@�C������t�H���_���Ɏg�p����ۂɎg�p����B
' = �ˑ�    �Ȃ�
' = ����    String.vbs
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateTime _
)
    On Error Resume Next
    Dim sDateStr
    sDateStr = _
        String(4 - Len(Year(sDateTime)),   "0") & Year(sDateTime)   & _
        String(2 - Len(Month(sDateTime)),  "0") & Month(sDateTime)  & _
        String(2 - Len(Day(sDateTime)),    "0") & Day(sDateTime)    & _
        "-" & _
        String(2 - Len(Hour(sDateTime)),   "0") & Hour(sDateTime)   & _
        String(2 - Len(Minute(sDateTime)), "0") & Minute(sDateTime) & _
        String(2 - Len(Second(sDateTime)), "0") & Second(sDateTime)
    If Err.Number = 0 Then
        ConvDate2String = sDateStr
    Else
        ConvDate2String = ""
    End If
    On Error Goto 0
End Function
    'Call Test_ConvDate2String()
    Private Sub Test_ConvDate2String()
        MsgBox  ConvDate2String(Now()) & vbNewLine & _
                ConvDate2String("2001/12/32 1:00:0")
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
            lPaddingTabNum = Application.WorksheetFunction.RoundUp(lPaddingLen / lTabWidth, 0)
            CalcPaddingTabWidth = True
        Else
            lPaddingTabNum = 0
            CalcPaddingTabWidth = False
        End If
    End If
End Function
    Call Test_CalcPaddingTabWidth()
    Private Sub Test_CalcPaddingTabWidth()
        '���v�e�X�g
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
                                                                                                  '
        Result = Result & vbNewLine & CalcPaddingWidth(0, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 8
        Result = Result & vbNewLine & CalcPaddingWidth(1, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(2, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 6
        Result = Result & vbNewLine & CalcPaddingWidth(3, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 5
        Result = Result & vbNewLine & CalcPaddingWidth(4, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & CalcPaddingWidth(5, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 3
        Result = Result & vbNewLine & CalcPaddingWidth(6, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 2
        Result = Result & vbNewLine & CalcPaddingWidth(7, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 1
        Result = Result & vbNewLine & CalcPaddingWidth(8, 1, 8, lPaddingLen)  & ":" & lPaddingLen ' 8
                                                                                                  '
        Result = Result & vbNewLine & CalcPaddingWidth(0, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(3, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 4
        Result = Result & vbNewLine & CalcPaddingWidth(4, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 3
        Result = Result & vbNewLine & CalcPaddingWidth(5, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 2
        Result = Result & vbNewLine & CalcPaddingWidth(6, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 1
        Result = Result & vbNewLine & CalcPaddingWidth(7, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 7
        Result = Result & vbNewLine & CalcPaddingWidth(8, 5, 7, lPaddingLen)  & ":" & lPaddingLen ' 6
                                                                                                  '
        Result = Result & vbNewLine & CalcPaddingWidth(1, 5, 0, lPaddingLen)  & ":" & lPaddingLen ' 0
        Result = Result & vbNewLine & CalcPaddingWidth(1, 5, -1, lPaddingLen) & ":" & lPaddingLen ' 0
        Result = Result & vbNewLine & CalcPaddingWidth(1, -1, 4, lPaddingLen) & ":" & lPaddingLen ' 0
        Result = Result & vbNewLine & CalcPaddingWidth(-1, 5, 4, lPaddingLen) & ":" & lPaddingLen ' 0
        
        MsgBox Result
    End Sub
