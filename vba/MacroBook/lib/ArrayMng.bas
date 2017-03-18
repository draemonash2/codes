Attribute VB_Name = "ArrayMng"
Option Explicit

' array manage library v1.0

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

