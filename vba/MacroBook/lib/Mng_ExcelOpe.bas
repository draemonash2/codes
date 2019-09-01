Attribute VB_Name = "Mng_ExcelOpe"
Option Explicit

' excel operation library v2.31

'************************************************************
'* �֐���`
'************************************************************
' ==================================================================
' = �T�v    �V�[�g�ꗗ�쐬
' = ����    �Ȃ�
' = �ߒl    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Sub CreateSheetList()
    Dim oSheet As Object
    Dim lRowIdx As Long
    Dim lColumnIdx As Long
 
    If MsgBox("�A�N�e�B�u�Z�����牺�ɃV�[�g���ꗗ���쐬���Ă������ł����H", vbYesNo + vbDefaultButton2) = vbNo Then
        'None
    Else
        lRowIdx = ActiveCell.Row
        lColumnIdx = ActiveCell.Column
 
        For Each oSheet In ActiveWorkbook.Sheets
            Cells(lRowIdx, lColumnIdx).Value = oSheet.Name
            lRowIdx = lRowIdx + 1
        Next oSheet
    End If
End Sub

' ==================================================================
' = �T�v    ���[�N�V�[�g��V�K�쐬
' =         �d���������[�N�V�[�g������ꍇ�A_1, _2 ...�ƘA�ԂɂȂ�B
' =         �Ăяo�����ɂ͍쐬�������[�N�V�[�g����Ԃ��B
' = ����    sSheetName  [in]    String  �쐬����V�[�g��
' = �ߒl                        String  �쐬�����V�[�g��
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Public Function CreateNewWorksheet( _
    ByVal sSheetName As String _
) As String
    Dim lShtIdx As Long
    
    lShtIdx = 0
    Dim bExistWorkSht As Boolean
    Do
        bExistWorkSht = ExistsWorksheet(sSheetName)
        If bExistWorkSht Then
            sSheetName = sSheetName & "_"
        Else
            lShtIdx = lShtIdx + 1 '�A�ԗp�̕ϐ�
        End If
    Loop While bExistWorkSht
    
    With ActiveWorkbook
        .Worksheets.Add(after:=.Worksheets(.Worksheets.Count)).Name = sSheetName
    End With
    CreateNewWorksheet = sSheetName
End Function

' ==================================================================
' = �T�v    �d������Worksheet���L�邩�`�F�b�N����B
' = ����    sTrgtShtName    [in]    String  �V�[�g��
' = �ߒl                            Boolean ���ݗL��
' = �ˑ�    �Ȃ�
' = ����    Mng_ExcelOpe.bas
' ==================================================================
Private Function ExistsWorksheet( _
    ByVal sTrgtShtName As String _
) As Boolean
    Dim lShtIdx As Long
    
    With ActiveWorkbook
        ExistsWorksheet = False
        For lShtIdx = 1 To .Worksheets.Count
            If .Worksheets(lShtIdx).Name = sTrgtShtName Then
                ExistsWorksheet = True
                Exit For
            End If
        Next
    End With
End Function

