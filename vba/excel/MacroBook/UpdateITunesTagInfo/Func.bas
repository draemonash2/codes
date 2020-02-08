Attribute VB_Name = "Func"
Option Explicit

' ==================================================================
' = �T�v    �w�肵���Q�͈̔͂��r���āA���S��v���ǂ����𔻒肷��
' = ����    rTrgtRange01        Range   [in]  ��r�Ώ۔͈͂P
' = ����    rTrgtRange02        Range   [in]  ��r�Ώ۔͈͂Q
' = ����    bCellPosCheckValid  Boolean [in]  �Z���ʒu�`�F�b�N�L��/����
' = �ߒl                        Boolean       ��r����
' = �o��    �ȉ��̂����ꂩ�𖞂����ꍇ�AFalse ��ԋp����
' =           �E�͈͓��̃Z�������s��v
' =           �E�͈͓��̍s�����s��v
' =           �E�͈͓��̗񐔂��s��v
' =           �E�͈͓��̊e�Z���̒l���s��v
' =           �E�͈͓��̊J�n�Z���Ɩ����Z���̃Z���ʒu���s��v
' ==================================================================
Public Function DiffRange( _
    ByRef rTrgtRange01 As Range, _
    ByRef rTrgtRange02 As Range, _
    Optional bCellPosCheckValid As Boolean = False _
) As Boolean
    DiffRange = True
    If rTrgtRange01.Count = rTrgtRange02.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    If rTrgtRange01.Rows.Count = rTrgtRange02.Rows.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    If rTrgtRange01.Columns.Count = rTrgtRange02.Columns.Count Then
        'Do Nothing
    Else
        DiffRange = False
        Exit Function
    End If
    Dim lIdx As Long
    For lIdx = 1 To rTrgtRange01.Count
        If rTrgtRange01(lIdx).Value = rTrgtRange02(lIdx).Value Then
            'Do Nothing
        Else
            DiffRange = False
            Exit Function
        End If
    Next lIdx
    If bCellPosCheckValid = True Then
        If rTrgtRange01(1).Row = rTrgtRange02(1).Row And _
           rTrgtRange01(1).Column = rTrgtRange02(1).Column And _
           rTrgtRange01(rTrgtRange01.Count).Row = rTrgtRange02(rTrgtRange02.Count).Row And _
           rTrgtRange01(rTrgtRange01.Count).Column = rTrgtRange02(rTrgtRange02.Count).Column Then
            'Do Nothing
        Else
            DiffRange = False
            Exit Function
        End If
    Else
        'Do Nothing
    End If
    
End Function
    Private Sub Test_DiffRange()
        Dim shDiff01 As Worksheet
        Dim shDiff02 As Worksheet
        Set shDiff01 = ThisWorkbook.Sheets("�^�O�ꗗ")
        Set shDiff02 = ThisWorkbook.Sheets("�^�O�ꗗ_�~���[")
        Debug.Print DiffRange( _
            shDiff01.Range( _
                shDiff01.Cells(4, 6), _
                shDiff01.Cells(4, 39) _
            ), _
            shDiff02.Range( _
                shDiff02.Cells(4, 6), _
                shDiff02.Cells(4, 39) _
            ) _
        )
    End Sub
