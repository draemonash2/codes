Attribute VB_Name = "Template"
Option Explicit

' templates v1.0

' =============================================================================
' = �T�v    ��
' = �o��    �Ȃ�
' = �ˑ�    XXX.bas/Xxxx()
' =         YYY.bas/Yyyy()
' = ����    ��
' =============================================================================
Private Sub TemplateSub()
    '�������ݒ� �������灥����
    Const sMACRO_NAME = "���}�N������"
    '�������ݒ� �����܂Ł�����
    
    Dim vCalcSetting As Variant
    Application.ScreenUpdating = False
    vCalcSetting = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim rFindResult As Range
    Dim sSrchKeyword As String
    
    With ThisWorkbook.Sheets("���V�[�g����")
        '*** ���O���� ***
        '�J�n�I���s,�񌟍�
        sSrchKeyword = "�������P�ꁚ"
        Set rFindResult = .Cells.Find(sSrchKeyword, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "��������܂���ł���", vbCritical, sMACRO_NAME
            MsgBox "�����𒆒f���܂�", vbCritical, sMACRO_NAME
            Exit Sub
        End If
        Dim lTitleRow As Long
        Dim lStrtRow As Long
        Dim lLastRow As Long
        Dim lStrtClm As Long
        Dim lLastClm As Long
        lTitleRow = rFindResult.Row
        lStrtRow = rFindResult.Row + 1
        lStrtClm = rFindResult.Column
        lLastRow = .Cells(.Rows.Count, lStrtClm).End(xlUp).Row
        lLastClm = .Cells(lTitleRow, .Columns.Count).End(xlToLeft).Column
        
        '*** �{���� ***
        Dim lRowIdx As Long
        For lRowIdx = lStrtRow To lLastRow
            '�������ɏ�����������
        Next lRowIdx
    End With
  
    Application.Calculation = vCalcSetting
    Application.ScreenUpdating = True
End Sub

' ==================================================================
' = �T�v    ��
' = ����    sAAAA           String   [in]   �����͕�����
' = �ߒl                    Boolean         ���߂�l��
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    XXX.bas
' ==================================================================
Private Function TemplateFunc()
    '�������ɏ�����������
End Function
    Private Sub Test_TemplateFunc()
        Dim bRet As Boolean
        Debug.Print "*** test start! ***"
        '�������ɏ�����������
        Debug.Print "*** test finished! ***"
    End Sub


