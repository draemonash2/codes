Attribute VB_Name = "Mng_Array"
Option Explicit

' array manage library v1.34

Public Enum E_INSERT_TYPE
    E_INSERT_TOP
    E_INSERT_MIDDLE
    E_INSERT_BOTTOM
End Enum

' ==================================================================
' = �T�v    String �z��ɑ΂��� Push ����B
' = ����    sPushStr  [in]  String      Push ���镶����
' = ����    asTrgtStr [Out] StrArray    Push �Ώ۔z��
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
' ==================================================================
Public Function PushToStrArray( _
    ByVal sPushStr As String, _
    ByRef asTrgtStr() As String _
)
    Dim lNewArrIdx As Long
 
    If Sgn(asTrgtStr) = 0 Then
        ReDim Preserve asTrgtStr(0)
        asTrgtStr(0) = sPushStr
    Else
        lNewArrIdx = UBound(asTrgtStr) + 1
        ReDim Preserve asTrgtStr(lNewArrIdx)
        asTrgtStr(lNewArrIdx) = sPushStr
    End If
End Function
 
' ==================================================================
' = �T�v    String �z��ɑ΂��� Pop ����B
' =         �������Ȃ��z�񂪎w�肳�ꂽ�ꍇ�A"" ��ԋp����B
' = ����    asSrcStr [In]  StrArray    Pop �Ώ۔z��
' = ����    sPopStr  [Out] String      Pop ����������
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
' ==================================================================
Public Function PopToStrArray( _
    ByRef asSrcStr() As String, _
    ByRef sPopStr As String _
)
    Dim lPrevArrIdx As Long
 
    If Sgn(asSrcStr) = 0 Then
        sPopStr = ""
    ElseIf UBound(asSrcStr) = 0 Then
        sPopStr = asSrcStr(0)
        ReDim asSrcStr(0)
    Else
        lPrevArrIdx = UBound(asSrcStr)
        sPopStr = asSrcStr(lPrevArrIdx)
        ReDim Preserve asSrcStr(lPrevArrIdx - 1)
    End If
End Function

' ==================================================================
' = �T�v    String �z��ɑ΂��Ďw��ʒu�ɔz���}������B
' = ����    eInsertType    [In]        Enum        �}����ʁi�擪/����/�����j
' = ����    lTrgtIdx       [In]        Long        �}���������v�f�ԍ�
' = ����    asTrgtStr()    [Out]       String      �}�������������z��
' = ����    asBaseStr()    [In,Out]    String()    �}���������z��A�}����̕����z��
' = �ߒl    �Ȃ�
' = �o��    
' =         ��P�j�z��ԍ� 0�`3 �̔z��ɑ΂��āA�}����ʁu�擪�v��
' =               �w�肵���ꍇ
' =                 0, 1, 2, 3 �� _, 0, 1, 2, 3
' =         ��Q�j�z��ԍ� 0�`3 �̔z��ɑ΂��āA�}����ʁu�����v��
' =               �w�肵���ꍇ
' =                 0, 1, 2, 3 �� 0, 1, 2, 3, _
' =         ��R�j�z��ԍ� 0�`3 �̔z��ɑ΂��āA�}����ʁu���ԁv�A
' =               lTrgtIdx = 2 ���w�肵���ꍇ
' =                 0, 1, 2, 3 �� 0, 1, _, 2, 3
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
' ==================================================================
Public Function InsertToStrArray( _
    ByRef eInsertType As E_INSERT_TYPE, _
    ByRef lTrgtIdx As Long, _
    ByRef asTrgtStr() As String, _
    ByRef asBaseStr() As String _
)
    'TODO�F�v����
End Function

' ==================================================================
' = �T�v    String �z��ɑ΂��Ďw��ʒu�ɔz���}������B�i�w��v�f�u�������j
' = ����    lTrgtIdx       [In]        Long        �u�����������v�f�ԍ�
' = ����    asRepArr()     [In]        String()    �u���������������z��
' = ����    asBaseArr()    [In,Out]    String()    �u�������������z��A�}����̕����z��
' = �ߒl                               Boolean     �u����������
' = �o��    
' =           ��P�j�z��A�iasBaseArr�v�f0�`4�j�ɑ΂��āA�z��B�iasRepArr�v�f0�`2)�A
' =                 lTrgtIdx = 2 ���w�肵���ꍇ
' =                       0     1     2     3     4
' =                   A = A[0], A[1], A[2], A[3], A[4]
' =                   ��
' =                       0     1     2     3     4     5     6
' =                   A = A[0], A[1], B[0], B[1], B[2], A[3], A[4]
' =           ��Q�j�z��A�iasBaseArr�v�f0�`4�j�ɑ΂��āA�z��B�iasRepArr��z��)�A
' =                 lTrgtIdx = 2 ���w�肵���ꍇ
' =                       0     1     2     3     4
' =                   A = A[0], A[1], A[2], A[3], A[4]
' =                   ��
' =                       0     1     2     3
' =                   A = A[0], A[1], A[3], A[4]
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
' ==================================================================
Public Function InsRepToStrArray( _
    ByRef lTrgtIdx As Long, _
    ByRef asRepArr() As String, _
    ByRef asBaseArr() As String _
) As Boolean
    Dim bIsError As Boolean
    bIsError = False
    
    '�����`�F�b�N
    'Debug.Assert Sgn(asRepArr) <> 0
    'Debug.Assert LBound(asRepArr) = 0
    Debug.Assert Sgn(asBaseArr) <> 0
    Debug.Assert LBound(asBaseArr) = 0
    Debug.Assert lTrgtIdx >= LBound(asBaseArr) And lTrgtIdx <= UBound(asBaseArr)
    
    Dim lBaseSrcIdx As Long
    Dim lBaseDstIdx As Long
    
    'asRepArr �����������z��̏ꍇ�A�v�f�ԍ� lTrgtIdx ���폜����
    If Sgn(asRepArr) = 0 Then
        For lBaseSrcIdx = (lTrgtIdx + 1) To UBound(asBaseArr)
            lBaseDstIdx = lBaseSrcIdx - 1
            asBaseArr(lBaseDstIdx) = asBaseArr(lBaseSrcIdx)
        Next lBaseSrcIdx
        ReDim Preserve asBaseArr(UBound(asBaseArr) - 1)
    
    'asRepArr ���������ςݔz��̏ꍇ�A�v�f�ԍ� lTrgtIdx �� asRepArr ��}������
    Else
        ReDim Preserve asBaseArr(UBound(asBaseArr) + UBound(asRepArr))
        '�ړ�
        For lBaseDstIdx = UBound(asBaseArr) To (lTrgtIdx + UBound(asRepArr) + 1) Step -1
            lBaseSrcIdx = lBaseDstIdx - UBound(asRepArr)
            asBaseArr(lBaseDstIdx) = asBaseArr(lBaseSrcIdx)
        Next lBaseDstIdx
        
        '�}��
        Dim lBaseIdx As Long
        Dim lRepIdx As Long
        For lBaseIdx = lTrgtIdx To (lTrgtIdx + UBound(asRepArr))
            'asRepArr�̗v�f���󕶎���̏ꍇ�AasBaseArr(lBaseIdx) = asRepArr(lRepIdx) �Ƃ����
            '�e�L�X�g�t�@�C���o�͎��Ɂi�Ȃ����j�G���[����������B
            '���̂��߁A�󕶎���������{�B
            If asRepArr(lRepIdx) = "" Then
                asBaseArr(lBaseIdx) = ""
            Else
                asBaseArr(lBaseIdx) = asRepArr(lRepIdx)
            End If
            lRepIdx = lRepIdx + 1
        Next lBaseIdx
    End If
End Function
    Private Function Test_InsRepToStrArray()
        Dim asBaseArr() As String
        Dim asRepArr() As String
        Dim asRepArr05() As String
        
        'asBaseArr(3),asRepArr(2)
        Call Test_InsRepToStrArraySub01(asBaseArr, asRepArr): Call InsRepToStrArray(0, asRepArr, asBaseArr)
        Call Test_InsRepToStrArraySub01(asBaseArr, asRepArr): Call InsRepToStrArray(2, asRepArr, asBaseArr)
        Call Test_InsRepToStrArraySub01(asBaseArr, asRepArr): Call InsRepToStrArray(3, asRepArr, asBaseArr)
       'Call Test_InsRepToStrArraySub01(asBaseArr, asRepArr): Call InsRepToStrArray(4, asRepArr, asBaseArr)
        
        'asBaseArr(3),asRepArr(0)
        Call Test_InsRepToStrArraySub02(asBaseArr, asRepArr): Call InsRepToStrArray(0, asRepArr, asBaseArr)
        Call Test_InsRepToStrArraySub02(asBaseArr, asRepArr): Call InsRepToStrArray(2, asRepArr, asBaseArr)
        Call Test_InsRepToStrArraySub02(asBaseArr, asRepArr): Call InsRepToStrArray(3, asRepArr, asBaseArr)
       'Call Test_InsRepToStrArraySub02(asBaseArr, asRepArr): Call InsRepToStrArray(4, asRepArr, asBaseArr)
        
        'asBaseArr(0),asRepArr(3)
        Call Test_InsRepToStrArraySub03(asBaseArr, asRepArr): Call InsRepToStrArray(0, asRepArr, asBaseArr)
       'Call Test_InsRepToStrArraySub03(asBaseArr, asRepArr): Call InsRepToStrArray(1, asRepArr, asBaseArr)
        
        'asBaseArr(0),asRepArr(0)
        Call Test_InsRepToStrArraySub04(asBaseArr, asRepArr): Call InsRepToStrArray(0, asRepArr, asBaseArr)
       'Call Test_InsRepToStrArraySub04(asBaseArr, asRepArr): Call InsRepToStrArray(1, asRepArr, asBaseArr)
        
        'asBaseArr(2),asRepArr(��������)
        Call Test_InsRepToStrArraySub05(asBaseArr, asRepArr05): Call InsRepToStrArray(0, asRepArr05, asBaseArr)
        Call Test_InsRepToStrArraySub05(asBaseArr, asRepArr05): Call InsRepToStrArray(1, asRepArr05, asBaseArr)
        Call Test_InsRepToStrArraySub05(asBaseArr, asRepArr05): Call InsRepToStrArray(2, asRepArr05, asBaseArr)
        
    End Function
        Private Function Test_InsRepToStrArraySub01( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(3)
            ReDim Preserve asRepArr(2)
            asBaseArr(0) = "0"
            asBaseArr(1) = "1"
            asBaseArr(2) = "2"
            asBaseArr(3) = "3"
            asRepArr(0) = "a"
            asRepArr(1) = "b"
            asRepArr(2) = "c"
        End Function
        Private Function Test_InsRepToStrArraySub02( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(3)
            ReDim Preserve asRepArr(0)
            asBaseArr(0) = "0"
            asBaseArr(1) = "1"
            asBaseArr(2) = "2"
            asBaseArr(3) = "3"
            asRepArr(0) = "a"
        End Function
        Private Function Test_InsRepToStrArraySub03( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(0)
            ReDim Preserve asRepArr(3)
            asBaseArr(0) = "0"
            asRepArr(0) = "a"
            asRepArr(1) = "b"
            asRepArr(2) = "c"
            asRepArr(3) = "d"
        End Function
        Private Function Test_InsRepToStrArraySub04( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(0)
            ReDim Preserve asRepArr(0)
            asBaseArr(0) = "0"
            asRepArr(0) = "a"
        End Function
        Private Function Test_InsRepToStrArraySub05( _
            ByRef asBaseArr() As String, _
            ByRef asRepArr() As String _
        )
            ReDim Preserve asBaseArr(2)
            'ReDim asRepArr(0)
            asBaseArr(0) = "0"
            asBaseArr(1) = "1"
            asBaseArr(2) = "2"
            'asRepArr(0) = "a"
        End Function

' ==================================================================
' = �T�v    �}����z��̒��g����L�[���[�h���������đ}���z��Œu������
' =         �i�e���v���[�g�t�@�C�������Ƀt�@�C���𐶐�����ۂɎg�p����j
' = ����    sKeyword        String      [in]    �L�[���[�h
' =         asRepArr()      String()    [in]    �}���z��
' =         asBaseArr()     String()    [out]   �}����z��
' =         bIsWholeMatch   Boolean     [in]    �L�[���[�h���S��v/������v�iTrue:���S��v�j
' = �ߒl                    Boolean             ��v����
' = �o��    asBaseArr �̒��ɓ���  �������܂܂�Ă���ꍇ�A
' =         �擪�� sKeyword �̂ݒu��������
' = �ˑ�    Mng_Array.bas/InsRepToStrArray()
' = ����    Mng_Array.bas
' ==================================================================
Public Function ReplaceArray( _
    ByVal sKeyword As String, _
    ByRef asRepArr() As String, _
    ByRef asBaseArr() As String, _
    Optional ByVal bIsWholeMatch As Boolean = True _
) As Boolean
    Dim lIdx As Long
    Dim bIsMatch As Boolean
    bIsMatch = False
    '���S��v
    If bIsWholeMatch = True Then
        For lIdx = LBound(asBaseArr) To UBound(asBaseArr)
            If asBaseArr(lIdx) = sKeyword Then
                Call InsRepToStrArray(lIdx, asRepArr, asBaseArr)
                bIsMatch = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lIdx
    '������v
    Else
        For lIdx = LBound(asBaseArr) To UBound(asBaseArr)
            If InStr(asBaseArr(lIdx), sKeyword) Then
                Call InsRepToStrArray(lIdx, asRepArr, asBaseArr)
                bIsMatch = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lIdx
    End If
    ReplaceArray = bIsMatch
End Function
    Private Sub Test_ReplaceArray()
        Dim asBaseFileLine() As String
        Dim asRepLine01() As String
        Dim asRepLine02() As String
        Dim asRepLine03() As String
        Dim asRepLine04() As String
        Dim asRepLine05() As String
        Dim sKeyword As String
        Dim bRet As Boolean
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
         '�f�X�N�g�b�v�t�H���_
        asBaseFileLine = InputTxtFile(objWshShell.SpecialFolders("Desktop") & "\" & "temp.vbs")
        
        ReDim Preserve asRepLine01(1)
        asRepLine01(0) = Chr(9) & "aaa"
        asRepLine01(1) = Chr(9) & "bbb"
        sKeyword = "'>>>�C���N���[�h<<<"
        bRet = ReplaceArray(sKeyword, asRepLine01, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "��������܂���ł���"
            Debug.Assert 0
        End If
        
        ReDim Preserve asRepLine02(0)
        'asRepLine02(0) = ""
        sKeyword = "'>>>�ϐ���`<<<"
        bRet = ReplaceArray(sKeyword, asRepLine02, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "��������܂���ł���"
            Debug.Assert 0
        End If
        
        ReDim Preserve asRepLine03(3)
        asRepLine03(0) = Chr(9) & "ccc"
        asRepLine03(1) = Chr(9) & "dddddd"
        asRepLine03(2) = ""
        asRepLine03(3) = Chr(9) & "e"
        sKeyword = "'>>>�{����<<<"
        bRet = ReplaceArray(sKeyword, asRepLine03, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "��������܂���ł���"
            Debug.Assert 0
        End If
        
        ReDim Preserve asRepLine04(0)
        asRepLine04(0) = "888888888888888"
        sKeyword = "'>>>�֐���`<<<"
        bRet = ReplaceArray(sKeyword, asRepLine04, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "��������܂���ł���"
            Debug.Assert 0
        End If
        
        'ReDim Preserve asRepLine05(0)
        'asRepLine05(0) = ""
        sKeyword = "'>>>�萔��`<<<"
        bRet = ReplaceArray(sKeyword, asRepLine05, asBaseFileLine)
        If bRet = True Then
            'Do Nothing
        Else
            MsgBox sKeyword & "��������܂���ł���"
            Debug.Assert 0
        End If
        
        Call OutputTxtFile(objWshShell.SpecialFolders("Desktop") & "\" & "temp2.vbs", asBaseFileLine)
    End Sub

' ==================================================================
' = �T�v    �Z���͈́iRange�^�j�𕶎���z��iString�z��^�j�ɕϊ�����B
' =         ��ɃZ���͈͂��e�L�X�g�t�@�C���ɏo�͂��鎞�Ɏg�p����B
' = ����    rCellsRange             Range   [in]  �Ώۂ̃Z���͈�
' = ����    asLine()                String  [out] ������ԊҌ�̃Z���͈�
' = ����    bIgnoreInvisibleCell    String  [in]  ��\���Z���������s��
' = ����    sDelimiter              String  [in]  ��؂蕶��
' = �ߒl    �Ȃ�
' = �o��    �񂪗ׂ荇�����Z�����m�͎w�肳�ꂽ��؂蕶���ŋ�؂���
' = �ˑ�    �Ȃ�
' = ����    Mng_Array.bas
' ==================================================================
Public Function ConvRange2Array( _
    ByRef rCellsRange As Range, _
    ByRef asLine() As String, _
    ByVal bIgnoreInvisibleCell As Boolean, _
    ByVal sDelimiter As String _
)
    Dim lLineIdx As Long
    lLineIdx = 0
    ReDim Preserve asLine(lLineIdx)
    
    Dim lRowIdx As Long
    For lRowIdx = 1 To rCellsRange.Rows.Count
        Dim lIgnoreCnt As Long
        lIgnoreCnt = 0
        Dim lClmIdx As Long
        For lClmIdx = 1 To rCellsRange.Columns.Count
            Dim sCurCellValue As String
            sCurCellValue = rCellsRange(lRowIdx, lClmIdx).Value
            '��\���Z���͖�������
            Dim bIsIgnoreCurExec As Boolean
            If bIgnoreInvisibleCell = True Then
                If rCellsRange(lRowIdx, lClmIdx).EntireRow.Hidden = True Or _
                   rCellsRange(lRowIdx, lClmIdx).EntireColumn.Hidden = True Then
                    bIsIgnoreCurExec = True
                Else
                    bIsIgnoreCurExec = False
                End If
            Else
                bIsIgnoreCurExec = False
            End If
            
            If bIsIgnoreCurExec = True Then
                lIgnoreCnt = lIgnoreCnt + 1
            Else
                If lClmIdx = 1 Then
                    asLine(lLineIdx) = sCurCellValue
                Else
                    asLine(lLineIdx) = asLine(lLineIdx) & sDelimiter & sCurCellValue
                End If
            End If
        Next lClmIdx
        If lIgnoreCnt = rCellsRange.Columns.Count Then '��\���s�͍s���Z���Ȃ�
            'Do Nothing
        Else
            If lRowIdx = rCellsRange.Rows.Count Then '�ŏI�s�͍s���Z���Ȃ�
                'Do Nothing
            Else
                lLineIdx = lLineIdx + 1
                ReDim Preserve asLine(lLineIdx)
            End If
        End If
    Next lRowIdx
End Function
    Private Sub Test_ConvRange2Array()
        '��
    End Sub

