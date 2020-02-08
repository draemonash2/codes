Attribute VB_Name = "Main_Common"
Option Explicit

'####################################################
'### タグ一覧シート
'####################################################
Public Const TAG_LIST_SHEET_NAME As String = "タグ一覧"

Public Const REF_CELL_SEARCH_KEY As String = "★"
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
'### ミラーシート
'####################################################
Public Const TAG_LIST_MIRROR_SHEET_NAME As String = "タグ一覧_ミラー"

'####################################################
'### ログシート
'####################################################
Public Const ERROR_LOG_SHEET_NAME As String = "エラーログ"
Public Const LOG_START_ROW  As Long = 2
Public Enum E_LOG_CLM
    LOG_CLM_DATETIME = 1
    LOG_CLM_RW
    LOG_CLM_FILEPATH
    LOG_CLM_ERRORMSG
End Enum
Public Const OUTPUT_SUCCESS_LOG_TO_ERROR_LOG As Boolean = False

 '環境ごとに異なる可能性がある。トラック名取得エラーが発生した場合、「Exec_GetDetailsOfGetDetailsOf()」の実行結果を
 'もとに以下のトラック名のインデックスを更新しておくこと。
Public Const FILE_DETAIL_INFO_TRACK_NAME_INDEX As Long = 21
Public Const FILE_DETAIL_INFO_TRACK_NAME_TITLE As String = "タイトル"

Public Function GetPreInfo()
    '行列取得
    Dim rFindResult As Range
    Dim sSrchKeyword As String
    Dim lSrchCellRow As Long
    Dim lSrchCellClm As Long
    With ThisWorkbook.Sheets(TAG_LIST_SHEET_NAME)
        '### 読書結果 ###
        sSrchKeyword = REF_CELL_SEARCH_KEY
        Set rFindResult = .Cells.Find(sSrchKeyword, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            MsgBox sSrchKeyword & "が見つかりませんでした"
            MsgBox "プログラムを終了します"
            End
        Else
            lSrchCellRow = rFindResult.Row
            lSrchCellClm = rFindResult.Column
        End If
        glRefStartRow = lSrchCellRow
        glRefStartClm = lSrchCellClm
    End With
End Function

