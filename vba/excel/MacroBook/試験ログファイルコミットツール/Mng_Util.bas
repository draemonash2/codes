Attribute VB_Name = "Mng_Util"
Option Explicit

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
    
    ' �����񕪊�
    vSplitStr = Split(sTrgtStr, sSrchStr)
    
    GetStrNum = UBound(vSplitStr)
End Function


