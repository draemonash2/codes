Option Explicit

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕������ԋp����B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ���o������
' = �o��    �Ȃ�
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
'   Private Sub Test()
'       Dim Result
'       Result = "[Result]"
'       Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "\" )   ' a.txt
'       Result = Result & vbNewLine & ExtractTailWord( "C:\test\a", "\" )       ' a
'       Result = Result & vbNewLine & ExtractTailWord( "C:\test\", "\" )        ' 
'       Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\" )         ' test
'       Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\\" )        ' C:\test
'       Result = Result & vbNewLine & ExtractTailWord( "a.txt", "\" )           ' a.txt
'       Result = Result & vbNewLine & ExtractTailWord( "", "\" )                ' 
'       Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "" )    ' C:\test\a.txt
'       MsgBox Result
'   End Sub
'   Call Test()

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕��������������B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ����������
' = �o��    �Ȃ�
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
'   Private Sub Test()
'       Dim Result
'       Result = "[Result]"
'       Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "\" )    ' C:\test
'       Result = Result & vbNewLine & RemoveTailWord( "C:\test\a", "\" )        ' C:\test
'       Result = Result & vbNewLine & RemoveTailWord( "C:\test\", "\" )         ' C:\test
'       Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\" )          ' C:
'       Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\\" )         ' C:\test
'       Result = Result & vbNewLine & RemoveTailWord( "", "\" )                 ' 
'       Result = Result & vbNewLine & RemoveTailWord( "a.txt", "\" )            ' a.txt�i�t�@�C�������ǂ����͔��f���Ȃ��j
'       Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "" )     ' C:\test\a.txt
'       MsgBox Result
'   End Sub
'   Call Test()

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�H���_�p�X
' = �o��    �Ȃ�
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath _
)
    GetDirPath = RemoveTailWord( sFilePath, "\" )
End Function
'   Private Sub Test()
'       'RemoveTailWord() �Ɠ����̃e�X�g�P�[�X�̂��߃e�X�g���Ȃ�
'   End Sub
'   Call Test()

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�@�C�����𒊏o����
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�@�C����
' = �o��    �Ȃ�
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath _
)
    GetFileName = ExtractTailWord( sFilePath, "\" )
End Function
'   Private Sub Test()
'       'ExtractTailWord() �Ɠ����̃e�X�g�P�[�X�̂��߃e�X�g���Ȃ�
'   End Sub
'   Call Test()

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����t�@�C�����i�g���q�Ȃ��j�𒊏o����
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �t�@�C�����i�g���q�Ȃ��j
' = �o��    �g���q���t�^����Ă��Ȃ��t�@�C�������݂���B���̂��߁A
' =         "." ���܂܂�Ă��Ȃ��ꍇ���������ԋp����B
' ==================================================================
Public Function GetFileBaseName( _
    ByVal sFilePath _
)
    GetFileBaseName = RemoveTailWord( ExtractTailWord( sFilePath, "\" ), "." )
End Function
'   Private Sub Test()
'       Dim Result
'       Result = "[Result]"
'       Result = Result & vbNewLine & GetFileBaseName( "C:\test\a.txt" )    ' a
'       Result = Result & vbNewLine & GetFileBaseName( "C:\test\a.t" )      ' a
'       Result = Result & vbNewLine & GetFileBaseName( "C:\test\a." )       ' a
'       Result = Result & vbNewLine & GetFileBaseName( "C:\test\a" )        ' a
'       Result = Result & vbNewLine & GetFileBaseName( "C:\test\" )         ' 
'       Result = Result & vbNewLine & GetFileBaseName( "C:\test" )          ' test
'       Result = Result & vbNewLine & GetFileBaseName( "C:" )               ' C:
'       Result = Result & vbNewLine & GetFileBaseName( "" )                 ' 
'       Result = Result & vbNewLine & GetFileBaseName( "a.txt" )            ' a
'       Result = Result & vbNewLine & GetFileBaseName( ".txt" )             ' 
'       Result = Result & vbNewLine & GetFileBaseName( "a." )               ' a
'       Result = Result & vbNewLine & GetFileBaseName( "." )                ' 
'       Result = Result & vbNewLine & GetFileBaseName( "a" )                ' a
'       MsgBox Result
'   End Sub
'   Call Test()

' ==================================================================
' = �T�v    �w�肳�ꂽ�t�@�C���p�X����g���q�𒊏o����
' = ����    sFilePath   String  [in]  �t�@�C���p�X
' = �ߒl                String        �g���q
' = �o��    "." ���܂܂�Ă��Ȃ��ꍇ�A�󕶎���ԋp����
' ==================================================================
Public Function GetFileExtName( _
    ByVal sFilePath _
)
    If InStr( sFilePath, "." ) > 0 Then
        GetFileExtName = ExtractTailWord( sFilePath, "." )
    Else
        GetFileExtName = ""
    End If
End Function
'   Private Sub Test()
'       Dim Result
'       Result = "[Result]"
'       Result = Result & vbNewLine & GetFileExtName( "C:\test\a.txt" ) ' txt
'       Result = Result & vbNewLine & GetFileExtName( "C:\test\a.t" )   ' t
'       Result = Result & vbNewLine & GetFileExtName( "C:\test\a." )    ' 
'       Result = Result & vbNewLine & GetFileExtName( "C:\test\a" )     ' 
'       Result = Result & vbNewLine & GetFileExtName( "C:\test\" )      ' 
'       Result = Result & vbNewLine & GetFileExtName( "C:\test" )       ' 
'       Result = Result & vbNewLine & GetFileExtName( "C:" )            ' 
'       Result = Result & vbNewLine & GetFileExtName( "" )              ' 
'       Result = Result & vbNewLine & GetFileExtName( "a.txt" )         ' txt
'       Result = Result & vbNewLine & GetFileExtName( ".txt" )          ' txt
'       Result = Result & vbNewLine & GetFileExtName( "a." )            ' 
'       Result = Result & vbNewLine & GetFileExtName( "." )             ' 
'       Result = Result & vbNewLine & GetFileExtName( "a" )             ' 
'       MsgBox Result
'   End Sub
'   Call Test()

' ==================================================================
' = �T�v    �w�肳�ꂽ������̕����񒷁i�o�C�g���j��ԋp����
' = ����    sInStr      String  [in]  ������
' = �ߒl                Long          �����񒷁i�o�C�g���j
' = �o��    �W���ŗp�ӂ���Ă��� LenB() �֐��́AUnicode �ɂ�����
' =         �o�C�g����ԋp���邽�ߔ��p�������Q�����Ƃ��ăJ�E���g����B
' =           �i��FLenB("�t�@�C���T�C�Y ") �� 16�j
' =         ���̂��߁A���p�������P�����Ƃ��ăJ�E���g����{�֐���p�ӁB
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
'   Private Sub Test()
'       Dim Result
'       Result = "[Result]"
'       Result = Result & vbNewLine & LenByte( "aaa" )      ' 3
'       Result = Result & vbNewLine & LenByte( "aaa " )     ' 4
'       Result = Result & vbNewLine & LenByte( "" )         ' 0
'       Result = Result & vbNewLine & LenByte( "������" )   ' 6
'       Result = Result & vbNewLine & LenByte( "������ " )  ' 7
'       Result = Result & vbNewLine & LenByte( "���� ��" )  ' 7
'       Result = Result & vbNewLine & LenByte( Chr(9) )     ' 1
'       Result = Result & vbNewLine & LenByte( Chr(10) )    ' 1
'       MsgBox Result
'   End Sub
'   Call Test()

