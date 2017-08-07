Attribute VB_Name = "StringMng"
Option Explicit

' string manage library v1.3

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
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�@�C���x�[�X���𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�@�C���x�[�X��
' = �o��    �E�g���q���Ȃ��ꍇ�A�󕶎���ԋp����
' =         �E�t�@�C�������w��\
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
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����g���q�𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �g���q
' = �o��    �E�g���q���Ȃ��ꍇ�A�󕶎���ԋp����
' =         �E�t�@�C�������w��\
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
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����w�肳�ꂽ�ꕔ�𒊏o����B
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = ����    lPartType   Long    [in]  ���o���
' =                                     1) �t�H���_�p�X
' =                                     2) �t�@�C����
' =                                     3) �t�@�C���x�[�X��
' =                                     4) �t�@�C���g���q
' = �ߒl                String        ���o�����ꕔ
' = �o��    �E���o��ʂ�����Ă���ꍇ�A�󕶎���ԋp����
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
' = �T�v    �����`����ϊ�����B�i��F2017/03/22 18:20:14 �� 20170322-182014�j
' = ����    sDateTime   String  [in]  �����iYYYY/MM/DD HH:MM:SS�j
' = �ߒl                String        �����iYYYYMMDD-HHMMSS�j
' = �o��    ��ɓ������t�@�C������t�H���_���Ɏg�p����ۂɎg�p����B
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
