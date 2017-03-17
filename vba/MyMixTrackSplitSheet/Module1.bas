Attribute VB_Name = "Module1"
Option Explicit

Const SHEET_NAME = "手入力"
Enum ROW
    ROW_TITLE = 2
    ROW_START = 3
End Enum
Const ALBUM_NAME_COLUMN = 4
Private TARGET_SERVICE_COLUMN_TABLE() As Long

Public Function ExportInformationsInit()
    ReDim Preserve TARGET_SERVICE_COLUMN_TABLE(2)
    TARGET_SERVICE_COLUMN_TABLE(0) = 14  'Audacity
    TARGET_SERVICE_COLUMN_TABLE(1) = 18  'MixCloud
    TARGET_SERVICE_COLUMN_TABLE(2) = 20  'SuperTagEditer
End Function

Public Sub ExportInformations()
    Call ExportInformationsInit
    
    Dim shTrgtSht As Worksheet
    Set shTrgtSht = ThisWorkbook.Sheets(SHEET_NAME)
    
    'フォルダ名指定
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFolderPath As String
    sOutputFolderPath = ShowFolderSelectDialog(objWshShell.SpecialFolders("Desktop"))

    'ファイル名指定
    Dim sOutputFileBaseName As String
    sOutputFileBaseName = InputBox( _
                                "ファイルベース名を入力してください", _
                                "test", _
                                shTrgtSht.Cells(ROW_START, ALBUM_NAME_COLUMN).Value _
                            )
    
    Dim lTrgtSrvcClmTblIdx As Long
    For lTrgtSrvcClmTblIdx = LBound(TARGET_SERVICE_COLUMN_TABLE) To UBound(TARGET_SERVICE_COLUMN_TABLE)
        Dim lClmIdx As Long
        lClmIdx = TARGET_SERVICE_COLUMN_TABLE(lTrgtSrvcClmTblIdx)
        
        '最終行判定
        Dim lLastRow As Long
        Dim lRowIdx As Long
        lRowIdx = ROW_START
        Do While 1
            If shTrgtSht.Cells(lRowIdx, lClmIdx).Value = "" Then
                Exit Do
            Else
                'Do Nothing
            End If
            lRowIdx = lRowIdx + 1
        Loop
        lLastRow = lRowIdx - 1
        If lLastRow < ROW_START Then
            Exit For
        Else
            'Do Nothing
        End If
        
        'ファイルパス作成
        Dim sTitleName As String
        sTitleName = shTrgtSht.Cells(ROW_TITLE, lClmIdx).Value
        Dim sOutputFilePath As String
        sOutputFilePath = sOutputFolderPath & "\" & sOutputFileBaseName & "_" & sTitleName & ".txt"
        
        'セル範囲取得
        Dim rTrgtRange As Range
        Set rTrgtRange = shTrgtSht.Range( _
                            shTrgtSht.Cells(ROW_START, lClmIdx), _
                            shTrgtSht.Cells(lLastRow, lClmIdx) _
                        )
        Dim asLine() As String
        Call ConvRange2Array(rTrgtRange, asLine(), True, Chr(9))
        
        'セル範囲出力
        Open sOutputFilePath For Output As #1
        Dim lLineIdx As Long
        For lLineIdx = LBound(asLine) To UBound(asLine)
            Print #1, asLine(lLineIdx)
        Next lLineIdx
        Close #1
    Next lTrgtSrvcClmTblIdx
    
    '本 Excel ファイルをコピー
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFile ThisWorkbook.FullName, sOutputFolderPath & "\" & sOutputFileBaseName & ".xlsm"
    
    MsgBox "エクスポート完了！"
End Sub

' ==================================================================
' = 概要    セル範囲（Range型）を文字列配列（String配列型）に変換する。
' =         主にセル範囲をテキストファイルに出力する時に使用する。
' = 引数    rCellsRange             Range   [in]  対象のセル範囲
' = 引数    asLine()                String  [out] 文字列返還後のセル範囲
' = 引数    bIsInvisibleCellIgnore  String  [in]  非表示セル無視実行可否
' = 引数    sDelimiter              String  [in]  区切り文字
' = 戻値    なし
' = 覚書    列が隣り合ったセル同士は指定された区切り文字で区切られる
' ==================================================================
Private Function ConvRange2Array( _
    ByRef rCellsRange As Range, _
    ByRef asLine() As String, _
    ByVal bIsInvisibleCellIgnore As Boolean, _
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
            '非表示セルは無視する
            Dim bIsIgnoreCurExec As Boolean
            If bIsInvisibleCellIgnore = True Then
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
        If lIgnoreCnt = rCellsRange.Columns.Count Then '非表示行は行加算しない
            'Do Nothing
        Else
            If lRowIdx = rCellsRange.Rows.Count Then '最終行は行加算しない
                'Do Nothing
            Else
                lLineIdx = lLineIdx + 1
                ReDim Preserve asLine(lLineIdx)
            End If
        End If
    Next lRowIdx
End Function

' ==================================================================
' = 概要    フォルダ選択ダイアログを表示する
' = 引数    sInitPath   String  [in]  デフォルトフォルダパス（省略可）
' = 戻値                String        フォルダ選択結果
' = 覚書    なし
' ==================================================================
Public Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "フォルダを選択してください"
    If sInitPath = "" Then
        'Do Nothing
    Else
        If Right(sInitPath, 1) = "\" Then
            fdDialog.InitialFileName = sInitPath
        Else
            fdDialog.InitialFileName = sInitPath & "\"
        End If
    End If
    
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ShowFolderSelectDialog = ""
    Else
        ShowFolderSelectDialog = fdDialog.SelectedItems.Item(1)
    End If
    
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFolderSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        MsgBox ShowFolderSelectDialog( _
                    objWshShell.SpecialFolders("Desktop") _
                )
    End Sub

