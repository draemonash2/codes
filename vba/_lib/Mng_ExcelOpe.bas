Attribute VB_Name = "Mng_ExcelOpe"
Option Explicit

' excel operation library v2.5

'************************************************************
'* 関数定義
'************************************************************
' ==================================================================
' = 概要    シート一覧作成
' = 引数    なし
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_ExcelOpe.bas
' ==================================================================
Sub CreateSheetList()
    Dim oSheet As Object
    Dim lRowIdx As Long
    Dim lColumnIdx As Long
 
    If MsgBox("アクティブセルから下にシート名一覧を作成してもいいですか？", vbYesNo + vbDefaultButton2) = vbNo Then
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
' = 概要    ワークシートを新規作成
' =         重複したワークシートがある場合、_1, _2 ...と連番になる。
' =         呼び出し側には作成したワークシート名を返す。
' = 引数    sSheetName  [in]    String  作成するシート名
' = 戻値                        String  作成したシート名
' = 依存    Mng_ExcelOpe.bas/ExistsWorksheet()
' = 所属    Mng_ExcelOpe.bas
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
            lShtIdx = lShtIdx + 1 '連番用の変数
        End If
    Loop While bExistWorkSht
    
    With ActiveWorkbook
        .Worksheets.Add(after:=.Worksheets(.Worksheets.Count)).Name = sSheetName
    End With
    CreateNewWorksheet = sSheetName
End Function

' ==================================================================
' = 概要    重複したWorksheetが有るかチェックする。
' = 引数    sTrgtShtName    [in]    String  シート名
' = 戻値                            Boolean 存在有無
' = 依存    なし
' = 所属    Mng_ExcelOpe.bas
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

' ==================================================================
' = 概要    指定したキーワードの近くのセル値を取得する
' = 引数    shTrgtSht       Worksheet   [in]    対象シート
' = 引数    sSearchKeyword  String      [in]    検索キーワード
' = 引数    lOffsetRow      Long        [in]    行オフセット
' = 引数    lOffsetClm      Long        [in]    列オフセット
' = 引数    sOutputValue    String      [out]   セル値
' = 戻値                    Boolean             取得結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_ExcelOpe.bas
' ==================================================================
Public Function GetNearCellValue( _
    ByRef shTrgtSht As Worksheet, _
    ByVal sSearchKeyword As String, _
    ByVal lOffsetRow As Long, _
    ByVal lOffsetClm As Long, _
    ByRef sOutputValue As String _
) As Boolean
    With shTrgtSht
        Dim rFindResult As Range
        Set rFindResult = .Cells.Find(sSearchKeyword, LookAt:=xlWhole)
        If rFindResult Is Nothing Then
            sOutputValue = ""
            GetNearCellValue = False
        Else
            If (rFindResult.Row + lOffsetRow) >= 1 And _
               (rFindResult.Column + lOffsetClm) >= 1 Then
                sOutputValue = .Cells( _
                                        rFindResult.Row + lOffsetRow, _
                                        rFindResult.Column + lOffsetClm _
                                    ).Value
                GetNearCellValue = True
            Else
                sOutputValue = ""
                GetNearCellValue = False
            End If
        End If
    End With
End Function
    Private Function Test_GetNearCellValue()
        Dim sSearchKeyword As String
        Dim lOffsetRow As Long
        Dim lOffsetClm As Long
        Dim sOutputValue As String
        Dim bRet As Boolean
        
        sSearchKeyword = "aaa"
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, 1, 1, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, 0, 1, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        sSearchKeyword = "bbb"
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -1, -1, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -1, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -2, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -3, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, -100, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, 0, -100, sOutputValue): Debug.Print bRet & " : " & sOutputValue
        sSearchKeyword = "ccc"
        bRet = GetNearCellValue(ActiveSheet, sSearchKeyword, 0, 0, sOutputValue): Debug.Print bRet & " : " & sOutputValue
    End Function

' ==================================================================
' = 概要    対象シートのセルを検索する。見つからない場合、処理を中断する。
' = 引数    shTrgtSht       Worksheet   [in]    検索対象シート
' = 引数    sFindKeyword    String      [in]    検索対象キーワード
' = 戻値                    Range               検索結果
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_ExcelOpe.bas
' ==================================================================
Public Function FindCell( _
    ByVal shTrgtSht As Worksheet, _
    ByVal sFindKeyword As String _
) As Range
    Set FindCell = shTrgtSht.Cells.Find(sFindKeyword, LookAt:=xlWhole)
    If FindCell Is Nothing Then
        MsgBox _
            "セルが見つからなかったため、処理を中断します。" & vbNewLine & _
            "　検索対象シート：" & shTrgtSht.Name & vbNewLine & _
            "　検索対象キーワード：" & sFindKeyword, _
            vbCritical
        End
    End If
End Function
    Private Function Test_FindCell()
        Dim rFindResult As Range
        Debug.Print "*** test start!"
        Set rFindResult = FindCell(ActiveSheet, "秀丸マクロ")
        Debug.Print "r" & rFindResult.Row & "c" & rFindResult.Column
        Set rFindResult = FindCell(ActiveSheet, "秀丸マク")
        Debug.Print "r" & rFindResult.Row & "c" & rFindResult.Column
        Debug.Print "*** test finish!"
    End Function



