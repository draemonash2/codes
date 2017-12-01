Attribute VB_Name = "Mng_Array"
Option Explicit

' array manage library v1.1

Public Enum E_INSERT_TYPE
    E_INSERT_TOP
    E_INSERT_MIDDLE
    E_INSERT_BOTTOM
End Enum

'===========================================================
'= �T�v�FString �z��ɑ΂��� Push ����B
'= �����FsPushStr  [in]  String      Push ���镶����
'= �����FasTrgtStr [Out] StrArray    Push �Ώ۔z��
'= �ߒl�F�Ȃ�
'===========================================================
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
 
'===========================================================
'= �T�v�FString �z��ɑ΂��� Pop ����B
'=       �������Ȃ��z�񂪎w�肳�ꂽ�ꍇ�A"" ��ԋp����B
'= �����FasSrcStr [In]  StrArray    Pop �Ώ۔z��
'= �����FsPopStr  [Out] String      Pop ����������
'= �ߒl�F�Ȃ�
'===========================================================
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

'===========================================================
'= �T�v�FString �z��ɑ΂��Ďw��ʒu�ɔz���}������B
'= �����FeInsertType    [In]        Enum        �}����ʁi�擪/����/�����j
'= �����FlTrgtIdx       [In]        Long        �}���������v�f�ԍ�
'= �����FasTrgtStr()    [Out]       String      �}�������������z��
'= �����FasBaseStr()    [In,Out]    String()    �}���������z��A�}����̕����z��
'= �ߒl�F�Ȃ�
'= �o���F
'=         ��P�j�z��ԍ� 0�`3 �̔z��ɑ΂��āA�}����ʁu�擪�v��
'=               �w�肵���ꍇ
'=                 0, 1, 2, 3 �� _, 0, 1, 2, 3
'=         ��Q�j�z��ԍ� 0�`3 �̔z��ɑ΂��āA�}����ʁu�����v��
'=               �w�肵���ꍇ
'=                 0, 1, 2, 3 �� 0, 1, 2, 3, _
'=         ��R�j�z��ԍ� 0�`3 �̔z��ɑ΂��āA�}����ʁu���ԁv�A
'=               lTrgtIdx = 2 ���w�肵���ꍇ
'=                 0, 1, 2, 3 �� 0, 1, _, 2, 3
'===========================================================
Public Function InsertToStrArray( _
    ByRef eInsertType As E_INSERT_TYPE, _
    ByRef lTrgtIdx As Long, _
    ByRef asTrgtStr() As String, _
    ByRef asBaseStr() As String _
)

End Function

'===========================================================
'= �T�v�FString �z��ɑ΂��Ďw��ʒu�ɔz���}������B�i�w��v�f�u�������j
'= �����FlTrgtIdx       [In]        Long        �u�����������v�f�ԍ�
'= �����FasRepArr()     [In]        String()    �u���������������z��
'= �����FasBaseArr()    [In,Out]    String()    �u�������������z��A�}����̕����z��
'= �ߒl�F                           Boolean     �u����������
'= �o���F
'=         ��P�j�z��A�i0�`4�j�ɑ΂��āA�z��B�i�v�f0�`2)�A
'=               lTrgtIdx = 2 ���w�肵���ꍇ
'=                     0     1     2     3     4
'=                 A = A[0], A[1], A[2], A[3], A[4]
'=                 ��
'=                     0     1     2     3     4     5     6
'=                 A = A[0], A[1], B[0], B[1], B[2], A[3], A[4]
'===========================================================
Public Function InsRepToStrArray( _
    ByRef lTrgtIdx As Long, _
    ByRef asRepArr() As String, _
    ByRef asBaseArr() As String _
) As Boolean
    Dim bIsError As Boolean
    bIsError = False
    
    '�����`�F�b�N
    Debug.Assert Sgn(asRepArr) <> 0
    Debug.Assert Sgn(asBaseArr) <> 0
    Debug.Assert lTrgtIdx >= LBound(asBaseArr) And lTrgtIdx <= UBound(asBaseArr)
    Debug.Assert LBound(asRepArr) = 0
    Debug.Assert LBound(asBaseArr) = 0
    
    ReDim Preserve asBaseArr(UBound(asBaseArr) + UBound(asRepArr))
    '�ړ�
    Dim lBaseSrcIdx As Long
    Dim lBaseDstIdx As Long
    For lBaseDstIdx = UBound(asBaseArr) To (lTrgtIdx + UBound(asRepArr) + 1) Step -1
        lBaseSrcIdx = lBaseDstIdx - UBound(asRepArr)
        asBaseArr(lBaseDstIdx) = asBaseArr(lBaseSrcIdx)
    Next lBaseDstIdx
    
    '�}��
    Dim lBaseIdx As Long
    Dim lRepIdx As Long
    For lBaseIdx = lTrgtIdx To (lTrgtIdx + UBound(asRepArr))
        asBaseArr(lBaseIdx) = asRepArr(lRepIdx)
        lRepIdx = lRepIdx + 1
    Next lBaseIdx
End Function
    Private Function Test_InsRepToStrArray()
        Dim asBaseArr() As String
        Dim asRepArr() As String
        
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
