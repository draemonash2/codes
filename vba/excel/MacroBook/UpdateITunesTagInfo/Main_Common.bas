Attribute VB_Name = "Main_Common"
Option Explicit

'####################################################
'### �^�O�ꗗ�V�[�g
'####################################################
Public Const TAG_LIST_SHEET_NAME As String = "�^�O�ꗗ"

Public Const REF_CELL_SEARCH_KEY As String = "��"
Public Enum E_ROW_OFFSET
    ROW_OFFSET_TITLE_01 = 0
    ROW_OFFSET_TITLE_02
    ROW_OFFSET_TAG_START
End Enum
Public Enum E_CLM_OFFSET
    CLM_OFFSET_TRACKINFO_EXECUTE_ENABLE
    CLM_OFFSET_TRACKINFO_FILEPATH
    CLM_OFFSET_TAGINFO_DIFF
    CLM_OFFSET_TAGINFO_TAGSTART
End Enum

Public glRefStartRow As Long
Public glRefStartClm As Long

'####################################################
'### �~���[�V�[�g
'####################################################
Public Const TAG_LIST_MIRROR_SHEET_NAME As String = "�^�O�ꗗ_�~���["

'####################################################
'### ���O�V�[�g
'####################################################
Public Const ERROR_LOG_SHEET_NAME As String = "�G���[���O"
Public Const LOG_START_ROW  As Long = 2
Public Enum E_LOG_CLM
    LOG_CLM_DATETIME = 1
    LOG_CLM_RW
    LOG_CLM_FILEPATH
    LOG_CLM_ERRORMSG
End Enum
Public Const OUTPUT_SUCCESS_LOG_TO_ERROR_LOG As Boolean = False

 '�����ƂɈقȂ�\��������B�g���b�N���擾�G���[�����������ꍇ�A�uExec_GetDetailsOfGetDetailsOf()�v�̎��s���ʂ�
 '���ƂɈȉ��̃g���b�N���̃C���f�b�N�X���X�V���Ă������ƁB
Public Const FILE_DETAIL_INFO_TRACK_NAME_INDEX As Long = 21
Public Const FILE_DETAIL_INFO_TRACK_NAME_TITLE As String = "�^�C�g��"

Public Function GetPreInfo()
    '�s��擾
    Dim rFindResult As Range
    Dim sSrchKeyword As String
    Dim lSrchCellRow As Long
    Dim lSrchCellClm As Long
    With ThisWorkbook.Sheets(TAG_LIST_SHEET_NAME)
        '### �Ǐ����� ###
        sSrchKeyword = REF_CELL_SEARCH_KEY
        Set rFindResult = .Cells.Find(sSrchKeyword, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "��������܂���ł���"
            MsgBox "�v���O�������I�����܂�"
            End
        Else
            lSrchCellRow = rFindResult.Row
            lSrchCellClm = rFindResult.Column
        End If
        glRefStartRow = lSrchCellRow
        glRefStartClm = lSrchCellClm
    End With
End Function

