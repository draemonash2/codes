Attribute VB_Name = "Mng_String"
Option Explicit

' string manage library v1.42

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
' = �ˑ�    Mng_String.bas/ExtractTailWord()
' =         Mng_String.bas/GetFileName()
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
' = �ˑ�    Mng_String.bas/RemoveTailWord()
' =         Mng_String.bas/GetFileName()
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
' =         Mng_String.bas/GetFileBase()
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
' = �T�v    �����`����ϊ�����B�i��F2017/03/22 18:20:14 �� 20170322-182014�j
' = ����    sDateTime   String  [in]  �����iYYYY/MM/DD HH:MM:SS�j
' = �ߒl                String        �����iYYYYMMDD-HHMMSS�j
' = �o��    ��ɓ������t�@�C������t�H���_���Ɏg�p����ۂɎg�p����B
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateTime As String _
) As String
    ConvDate2String = Year(sDateTime) & _
                     String(2 - Len(Month(sDateTime)), "0") & Month(sDateTime) & _
                     String(2 - Len(Day(sDateTime)), "0") & Day(sDateTime) & _
                     "-" & _
                     String(2 - Len(Hour(sDateTime)), "0") & Hour(sDateTime) & _
                     String(2 - Len(Minute(sDateTime)), "0") & Minute(sDateTime) & _
                     String(2 - Len(Second(sDateTime)), "0") & Second(sDateTime)
End Function
    Private Sub Test_ConvDate2String()
        Debug.Print ConvDate2String(Now())
    End Sub


' ==================================================================
' = �T�v    ���� �^�ϊ�(String��Long)
' = ����    sNum            String  [in]  ����(String�^)
' = �ߒl                    Long          ����(Long�^)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function NumConvStr2Lng( _
    ByVal sNum As String _
) As Long
    NumConvStr2Lng = Asc(sNum) + 30913
End Function
    Private Sub Test_NumConvStr2Lng()
        '��
    End Sub

' ==================================================================
' = �T�v    ���� �^�ϊ�(Long��String)
' = ����    sNum            Long    [in]    ����(Long�^)
' = �ߒl                    String          ����(String�^)
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function NumConvLng2Str( _
    ByVal lNum As Long _
) As String
    NumConvLng2Str = Chr(lNum - 30913)
End Function
    Private Sub Test_NumConvLng2Str()
        '��
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
' = �T�v    ���K�\���������s���iVba�}�N���֐��p�j
' = ����    sTargetStr      String  [in]  �����Ώە�����
' = ����    sSearchPattern  String  [in]  �����p�^�[��
' = ����    oMatchResult    Object  [out] ��������
' = �ߒl                    Boolean       �q�b�g�L��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_String.bas
' ==================================================================
Public Function ExecRegExp( _
    ByVal sTargetStr As String, _
    ByVal sSearchPattern As String, _
    ByRef oMatchResult As Object, _
    Optional ByVal bIgnoreCase As Boolean = True, _
    Optional ByVal bGlobal As Boolean = True _
) As Boolean
    Dim oRegExp As Object
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.IgnoreCase = bIgnoreCase
    oRegExp.Global = bGlobal
    oRegExp.Pattern = sSearchPattern
    Set oMatchResult = oRegExp.Execute(sTargetStr)
    If oMatchResult.Count = 0 Then
        ExecRegExp = False
    Else
        ExecRegExp = True
    End If
End Function
    Private Sub Test_ExecRegExp()
        Dim sTargetStr As String
        Dim oMatchResult As Object
        sTargetStr = "void TestFunc(int arg1, char arg2);"
        Debug.Print "*** test start! ***"
        Debug.Print ExecRegExp(sTargetStr, " \w+\(", oMatchResult)
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

