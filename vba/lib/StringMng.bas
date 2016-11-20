Attribute VB_Name = "StringMng"
Option Explicit

' string manage library v1.0

' ==================================================================
' = �T�v    �t���p�X���� "�t�@�C����" �𒊏o����
' = ����    sFullPath   String  [in]  �t���p�X
' = �ߒl                Variant       �t�@�C����
' = �o��    �Ȃ�
' ==================================================================
Public Function GetFileName( _
    ByVal sFilePath As String _
) As Variant
    Dim asSplitWord() As String
    If InStr(sFilePath, "\") > 0 Then
        asSplitWord = Split(sFilePath, "\")
        GetFileName = asSplitWord(UBound(asSplitWord))
    Else
        GetFileName = CVErr(xlErrNA)  '�G���[�l
    End If
End Function
 
' ==================================================================x
' = �T�v    �t���p�X���� "�f�B���N�g���p�X" �𒊏o����
' = ����    sFullPath   String  [in]  �t���p�X
' = �ߒl                Variant       �f�B���N�g���p�X
' = �o��    �Ȃ�
' ==================================================================
Public Function GetDirPath( _
    ByVal sFilePath As String _
) As Variant
    Dim asSplitWord() As String
    Dim sFileName As String
    If InStr(sFilePath, "\") > 0 Then
        asSplitWord = Split(sFilePath, "\")
        sFileName = asSplitWord(UBound(asSplitWord))
        GetDirPath = Replace( _
                                sFilePath, _
                                "\" & sFileName, _
                                "" _
                            )
    Else
        GetDirPath = CVErr(xlErrNA)  '�G���[�l
    End If
End Function
 
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

