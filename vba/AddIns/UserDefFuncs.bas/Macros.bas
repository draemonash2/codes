Attribute VB_Name = "Macros"
Option Explicit

' user define macros v2.2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' =============================================================================
' =  <<マクロ一覧>>
' =    ・選択範囲内で中央                   選択セルに対して「選択範囲内で中央」を実行する
' =
' =    ・ダブルクォートを除いてセルコピー   ダブルクオーテーションなしでセルコピーする
' =    ・選択範囲をファイルエクスポート     選択範囲をファイルとしてエクスポートする。
' =    ・選択範囲をコマンド実行             選択範囲内のコマンドを実行する。
' =
' =    ・全シート名をコピー                 ブック内のシート名を全てコピーする
' =    ・シート表示非表示を切り替え         シート表示/非表示を切り替える
' =    ・シート並べ替え作業用シートを作成   シート並べ替え作業用シート作成
' =
' =    ・セル内の丸数字をデクリメント       ②～⑮を指定して、指定番号以降をインクリメントする
' =    ・セル内の丸数字をインクリメント     ①～⑭を指定して、指定番号以降をデクリメントする
' =
' =    ・ツリーをグループ化                 ツリーグループ化する
' =    ・ハイパーリンク一括オープン         選択した範囲のハイパーリンクを一括で開く
' =
' =    ・フォント色をトグル                 フォント色を「赤」⇔「自動」でトグルする
' =    ・背景色をトグル                     背景色を「黄」⇔「背景色なし」でトグルする
' =
' =    ・オートフィル実行                   オートフィルを実行する
' =============================================================================

'******************************************************************************
'* 定数定義
'******************************************************************************
'=== 以下、セル内の丸数字をデクリメント()/セル内の丸数字をインクリメント() 用定義 ===
Const NUM_MAX = 15
Const NUM_MIN = 1

'=== 以下、シート並べ替え作業用シートを作成() 用定義 ===
Private Const WORK_SHEET_NAME = "シート並べ替え作業用"

Enum E_ROW
    ROW_BTN = 2
    ROW_TEXT_1 = 4
    ROW_TEXT_2
    ROW_SHT_NAME_TITLE = 7
    ROW_SHT_NAME_STRT
End Enum

Enum E_CLM
    CLM_BTN = 2
    CLM_SHT_NAME = 2
End Enum

Private Type T_SHORTCUT_KEY
    sKey As String
    sMacroName As String
End Type
Dim gtShortcutKey() As T_SHORTCUT_KEY

' *****************************************************************************
' * ショートカットキー定義
' *****************************************************************************
Private Function InitUserDefShortcut()
    ' <<ショートカットキー追加方法>>
    '   (1) 以下の追加先に「AddUserDefShortcut()」呼び出しを追加する。
    '       第一引数にはショートカットキー、第二引数にマクロ名を指定する。
    '       ショートカットキーは Ctrl や Shift などと組み合わせて指定できる。
    '         Shift：+
    '         Ctrl ：^
    '         Alt  ：%
    '       詳細は以下 URL 参照。
    '         https://msdn.microsoft.com/ja-jp/library/office/ff197461.aspx
    '   (2) マクロ「ユーザー定義ショートカットキーを設定()」を実行する。
    '
    ' <<ショートカットキー解除方法>>
    '   (1) マクロ「ユーザー定義ショートカットキーを解除()」を実行する。
    
    '▼▼▼ 追加先 ▼▼▼
'   Call AddUserDefShortcut("   ", "選択範囲内で中央")

    Call AddUserDefShortcut("^+c", "ダブルクォートを除いてセルコピー")
'   Call AddUserDefShortcut("   ", "選択範囲をファイルエクスポート")
'   Call AddUserDefShortcut("   ", "選択範囲をコマンド実行")

'   Call AddUserDefShortcut("   ", "全シート名をコピー")
'   Call AddUserDefShortcut("   ", "シート表示非表示を切り替え")
'   Call AddUserDefShortcut("   ", "シート並べ替え作業用シートを作成")

'   Call AddUserDefShortcut("   ", "セル内の丸数字をデクリメント")
'   Call AddUserDefShortcut("   ", "セル内の丸数字をインクリメント")

'   Call AddUserDefShortcut("   ", "ツリーをグループ化")
'   Call AddUserDefShortcut("   ", "ハイパーリンク一括オープン")

'   Call AddUserDefShortcut("   ", "フォント色をトグル")
'   Call AddUserDefShortcut("   ", "背景色をトグル")
    
    Call AddUserDefShortcut("%^+{DOWN}", "'オートフィル実行(""Down"")'")
    Call AddUserDefShortcut("%^+{UP}", "'オートフィル実行(""Up"")'")
    Call AddUserDefShortcut("%^+{RIGHT}", "'オートフィル実行(""Right"")'")
    Call AddUserDefShortcut("%^+{LEFT}", "'オートフィル実行(""Left"")'")
    '▲▲▲ 追加先 ▲▲▲
End Function

Public Sub ユーザー定義ショートカットキーを設定()
    Call InitUserDefShortcut
    Call EnableUserDefShortcut
End Sub

Public Sub ユーザー定義ショートカットキーを解除()
    Call InitUserDefShortcut
    Call DisableUserDefShortcut
End Sub

' *****************************************************************************
' * マクロ定義
' *****************************************************************************
' =============================================================================
' = 概要：選択セルに対して「選択範囲内で中央」を実行する
' =============================================================================
Public Sub 選択範囲内で中央()
    Selection.HorizontalAlignment = xlCenterAcrossSelection
End Sub

' =============================================================================
' = 概要：①～⑭を指定して、指定番号以降をデクリメントする
' =============================================================================
Public Sub セル内の丸数字をデクリメント()
    Dim lTrgtNum As Long
    Dim sTrgtNum As String
    Dim lLoopCnt As Long
    
    sTrgtNum = InputBox("デクリメントします。" & vbNewLine & "開始番号を入力してください。（②～⑮）", "番号入力", "")
    
    '入力値チェック
    If sTrgtNum = "" Then: MsgBox "入力値エラー！": Exit Sub
    lTrgtNum = NumConvStr2Lng(sTrgtNum)
    If (lTrgtNum > NUM_MAX Or NUM_MIN + 1 > lTrgtNum) Then: MsgBox "入力値エラー！": Exit Sub
    
    '本処理
    For lLoopCnt = lTrgtNum To NUM_MAX
        Selection.Replace _
            what:=NumConvLng2Str(lLoopCnt), _
            replacement:=NumConvLng2Str(lLoopCnt - 1)
    Next lLoopCnt
    MsgBox "置換完了！"
End Sub

' =============================================================================
' = 概要：②～⑮を指定して、指定番号以降をインクリメントする
' =============================================================================
Public Sub セル内の丸数字をインクリメント()
    Dim lTrgtNum As Long
    Dim sTrgtNum As String
    Dim lLoopCnt As Long
    
    sTrgtNum = InputBox("インクリメントします。" & vbNewLine & "開始番号を入力してください。（①～⑭）", "番号入力", "")
    
    '入力値チェック
    If sTrgtNum = "" Then: MsgBox "入力値エラー！": Exit Sub
    lTrgtNum = NumConvStr2Lng(sTrgtNum)
    If (lTrgtNum > NUM_MAX - 1 Or NUM_MIN > lTrgtNum) Then: MsgBox "入力値エラー！": Exit Sub
    
    '本処理
    For lLoopCnt = NUM_MAX - 1 To lTrgtNum Step -1
        Selection.Replace _
            what:=NumConvLng2Str(lLoopCnt), _
            replacement:=NumConvLng2Str(lLoopCnt + 1)
    Next lLoopCnt
    MsgBox "置換完了！"
End Sub

' =============================================================================
' = 概要：ブック内のシート名を全てコピーする
' = 備考：本マクロがエラーとなる場合、以下のいずれかを実施すること。
' =       ・ツール->参照設定 にて「Microsoft Forms 2.0 Object Library」を選択
' =       ・ツール->参照設定 内の「参照」にて system32 内の「FM20.DLL」を選択
' =============================================================================
Public Sub 全シート名をコピー()
    Dim oSheet As Object
    Dim sSheetNames As String
    Dim doDataObj As New DataObject
    
    For Each oSheet In ActiveWorkbook.Sheets
        If sSheetNames = "" Then
            sSheetNames = oSheet.Name
        Else
            sSheetNames = sSheetNames + vbNewLine + oSheet.Name
        End If
    Next oSheet
    
    doDataObj.SetText sSheetNames
    doDataObj.PutInClipboard
    
    MsgBox "ブック内のシート名を全てコピーしました"
End Sub

' =============================================================================
' = 概要：シート表示/非表示を切り替える
' =============================================================================
Public Sub シート表示非表示を切り替え()
    SheetVisibleSetting.Show
End Sub

' =============================================================================
' = 概要：ダブルクオーテーションなしでセルコピーする
' =       非表示セルは無視する。複数範囲は未対応。
' =============================================================================
Public Sub ダブルクォートを除いてセルコピー()
    '*** 非表示セル出力判定 ***
    Dim bIsInvisibleCellIgnore As Boolean
    'ユーザー操作を単純化するため、デフォルトで「非表示セル無視」としておく
    bIsInvisibleCellIgnore = True
'    vAnswer = MsgBox("非表示セルを無視しますか？", vbYesNoCancel)
'    If vAnswer = vbYes Then
'        bIsInvisibleCellIgnore = True
'    ElseIf vAnswer = vbNo Then
'        bIsInvisibleCellIgnore = False
'    Else
'        MsgBox "処理を中断します"
'        End
'    End If
    
    '*** 区切り文字判定 ***
    Dim sDelimiter As String
    'ユーザー操作を単純化するため、列間の区切り文字はデフォルトで「タブ文字」固定としておく
    sDelimiter = Chr(9)
    
    '*** セル範囲をString()型へ変換 ***
    Dim asLine() As String
    Call ConvRange2Array( _
        Selection, _
        asLine, _
        bIsInvisibleCellIgnore, _
        sDelimiter _
    )
    
    'String()型を順次クリップボードにコピー
    Dim sBuf As String
    sBuf = ""
    Dim lLineIdx As Long
    For lLineIdx = LBound(asLine) To UBound(asLine)
        If lLineIdx = LBound(asLine) Then
            sBuf = asLine(lLineIdx)
        Else
            sBuf = sBuf & vbNewLine & asLine(lLineIdx)
        End If
    Next lLineIdx
    Call CopyText(sBuf)
    
    'フィードバック
    Application.StatusBar = "■■■■■■■■ コピー完了！ ■■■■■■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' =============================================================================
' = 概要：選択範囲をファイルとしてエクスポートする。
' =       隣り合った列のセルにはタブ文字を挿入して出力する。
' =============================================================================
Public Sub 選択範囲をファイルエクスポート()
    Const TEMP_FILE_NAME As String = "ExportCellRange.tmp"
    
    '*** セル選択判定 ***
    If Selection.Count = 0 Then
        MsgBox "セルが選択されていません"
        MsgBox "処理を中断します"
        End
    End If
    
    '*** Tempファイル読出し ***
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
        Dim sTmpPath As String
    sTmpPath = objWshShell.SpecialFolders("Templates") & "\" & TEMP_FILE_NAME
    Dim sDirPathOld As String
    Dim sFileNameOld As String
    If objFSO.FileExists(sTmpPath) Then
        Open sTmpPath For Input As #1
        Line Input #1, sDirPathOld
        Line Input #1, sFileNameOld
        Close #1
    Else
        sDirPathOld = objWshShell.SpecialFolders("Desktop")
        sFileNameOld = "export"
    End If
    
    '*** フォルダパス入力 ***
    Dim sOutputDirPath As String
    sOutputDirPath = ShowFolderSelectDialog(sDirPathOld)
    If sOutputDirPath = "" Then
        MsgBox "無効なフォルダを指定もしくはフォルダが選択されませんでした。"
        MsgBox "処理を中断します。"
        End
    Else
        'Do Nothing
    End If
    
    '*** ファイル名入力 ***
    Dim sOutputFileName As String
    sOutputFileName = InputBox("ファイル名を入力してください。（拡張子なし）", "ファイル名入力", sFileNameOld)
    
    '*** ファイル名作成 ***
    Dim sOutputFilePath As String
    sOutputFilePath = sOutputDirPath & "\" & sOutputFileName & ".txt"
    
    '*** ファイル上書き判定 ***
    If objFSO.FileExists(sOutputFilePath) Then
        Dim vAnswer As Variant
        vAnswer = MsgBox("ファイルが存在します。上書きしますか？", vbOKCancel)
        If vAnswer = vbOK Then
            'Do Nothing
        Else
            MsgBox "処理を中断します。"
            End
        End If
    Else
        'Do Nothing
    End If
    
    '*** 非表示セル出力判定 ***
    Dim bIsInvisibleCellIgnore As Boolean
    bIsInvisibleCellIgnore = True 'ユーザー操作を単純化するため、デフォルトで「非表示セル無視」としておく
'    vAnswer = MsgBox("非表示セルを無視しますか？", vbYesNoCancel)
'    If vAnswer = vbYes Then
'        bIsInvisibleCellIgnore = True
'    ElseIf vAnswer = vbNo Then
'        bIsInvisibleCellIgnore = False
'    Else
'        MsgBox "処理を中断します"
'        End
'    End If
    
    '*** ファイル出力処理 ***
    Dim sDelimiter As String
    sDelimiter = Chr(9) '列間の区切り文字は「タブ文字」固定
    
    'Range型からString()型へ変換
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIsInvisibleCellIgnore, _
                sDelimiter _
            )
    
    On Error Resume Next
    Open sOutputFilePath For Output As #1
    If Err.Number = 0 Then
        'Do Nothing
    Else
        MsgBox "無効なファイルパスが指定されました" & Err.Description
        MsgBox "処理を中断します。"
        End
    End If
    On Error GoTo 0
    Dim lLineIdx As Long
    For lLineIdx = LBound(asRange) To UBound(asRange)
        Print #1, asRange(lLineIdx)
    Next lLineIdx
    Close #1
    
    '*** Tempファイル書き出し ***
    Open sTmpPath For Output As #1
    Print #1, sOutputDirPath
    Print #1, sOutputFileName
    Close #1
    
    MsgBox "出力完了！"
    
    '*** 出力ファイルを開く ***
    If Left(sOutputFilePath, 1) = "" Then
        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
    Else
        'Do Nothing
    End If
    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = 概要：選択範囲内のコマンドを実行する。
' =       単一列選択時のみ有効。
' =============================================================================
Public Sub 選択範囲をコマンド実行()
    '*** セル選択判定 ***
    If Selection.Count = 0 Then
        MsgBox "セルが選択されていません"
        MsgBox "処理を中断します"
        End
    End If
    
    '*** 範囲チェック ***
    If Selection.Columns.Count = 1 Then
        'Do Nothing
    Else
        MsgBox "単一列のみ選択してください"
        MsgBox "処理を中断します"
        End
    End If
    
    '*** 非表示セル出力判定 ***
    Dim bIsInvisibleCellIgnore As Boolean
    bIsInvisibleCellIgnore = True 'ユーザー操作を単純化するため、デフォルトで「非表示セル無視」としておく
'    vAnswer = MsgBox("非表示セルを無視しますか？", vbYesNoCancel)
'    If vAnswer = vbYes Then
'        bIsInvisibleCellIgnore = True
'    ElseIf vAnswer = vbNo Then
'        bIsInvisibleCellIgnore = False
'    Else
'        MsgBox "処理を中断します"
'        End
'    End If
    
    'Range型からString()型へ変換
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIsInvisibleCellIgnore, _
                "" _
            )
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\redirect.log"
    
    '*** コマンド実行 ***
    Open sOutputFilePath For Append As #1
    Print #1, "****************************************************"
    Print #1, Now()
    Print #1, "****************************************************"
    Dim lLineIdx As Long
    For lLineIdx = LBound(asRange) To UBound(asRange)
        Print #1, asRange(lLineIdx)
        Print #1, ExecDosCmd(asRange(lLineIdx))
    Next lLineIdx
    Print #1, ""
    Close #1
    
    MsgBox "実行完了！"
    
    '*** 出力ファイルを開く ***
    If Left(sOutputFilePath, 1) = "" Then
        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
    Else
        'Do Nothing
    End If
    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = 概要：行をツリー構造にしてグループ化
' = Usage：ツリーグループ化したい範囲を選択し、マクロ「ツリーをグループ化」を実行する
' =============================================================================
Public Sub ツリーをグループ化()
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lStrtClm As Long
    Dim lLastClm As Long
    
    'グループ化設定変更
    ActiveSheet.Outline.SummaryRow = xlAbove
    
    lStrtRow = Selection(1).Row
    lLastRow = Selection(Selection.Count).Row
    lStrtClm = Selection(1).Column
    lLastClm = Selection(Selection.Count).Column
    
    'グループ化
    Call TreeGroupSub( _
       ActiveSheet, _
       lStrtRow, _
       lLastRow, _
       lStrtClm, _
       lLastClm _
    )
End Sub

' =============================================================================
' = 概要：シートを並び替える。
' =       本処理を実行すると、シート並べ替え作業用シートを作成する。
' =============================================================================
'並べ替えシート 作業用シート作成
Public Sub シート並べ替え作業用シートを作成()
    Dim lShtIdx As Long
    Dim asShtName() As String
    Dim shWorkSht As Worksheet
    Dim bExistWorkSht As Boolean
    Dim lRowIdx As Long
    Dim lClmIdx As Long
    Dim lArrIdx As Long
    
    With ActiveWorkbook
        Application.ScreenUpdating = False

        ' === シート情報取得 ===
        ReDim Preserve asShtName(.Worksheets.Count - 1)
        For lShtIdx = 1 To .Worksheets.Count
            asShtName(lShtIdx - 1) = .Sheets(lShtIdx).Name
        Next lShtIdx

        ' === 作業用シート作成 ===
        bExistWorkSht = False
        For lShtIdx = 1 To .Worksheets.Count
            If .Sheets(lShtIdx).Name = WORK_SHEET_NAME Then
                bExistWorkSht = True
                Exit For
            Else
                'Do Nothing
            End If
        Next lShtIdx
        If bExistWorkSht = True Then
            MsgBox "既に「" & WORK_SHEET_NAME & "」シートが作成されています。"
            MsgBox "処理を続けたい場合は、シートを削除してください。"
            MsgBox "処理を中断します。"
            End
        Else
            Set shWorkSht = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            shWorkSht.Name = WORK_SHEET_NAME
        End If

        'シート情報書き込み
        shWorkSht.Cells(ROW_TEXT_1, CLM_SHT_NAME).Value = "希望通りにシート名を並べ替えてください。（上から順に並べ替えます）"
        shWorkSht.Cells(ROW_TEXT_2, CLM_SHT_NAME).Value = "並べ替えが終わったら、「並べ替え実行！！」ボタンを押してください。"
        shWorkSht.Cells(ROW_SHT_NAME_TITLE, CLM_SHT_NAME).Value = "シート名"
        lArrIdx = 0
        For lRowIdx = ROW_SHT_NAME_STRT To ROW_SHT_NAME_STRT + UBound(asShtName)
            shWorkSht.Cells(lRowIdx, CLM_SHT_NAME).NumberFormatLocal = "@"
            shWorkSht.Cells(lRowIdx, CLM_SHT_NAME).Value = asShtName(lArrIdx)
            lArrIdx = lArrIdx + 1
        Next lRowIdx

        'ボタン追加
        With shWorkSht.Buttons.Add( _
            shWorkSht.Cells(ROW_BTN, CLM_BTN).Left, _
            shWorkSht.Cells(ROW_BTN, CLM_BTN).Top, _
            shWorkSht.Cells(ROW_BTN, CLM_BTN).Width, _
            shWorkSht.Cells(ROW_BTN, CLM_BTN).Height _
        )
            .OnAction = "SortSheetPost"
            .Characters.Text = "並べ替え実行！！"
        End With

        '書式設定
        With ActiveSheet
            .Cells(ROW_SHT_NAME_TITLE, CLM_SHT_NAME).Interior.ColorIndex = 34
            .Cells(ROW_BTN, CLM_BTN).RowHeight = 30
            .Cells(ROW_BTN, CLM_BTN).ColumnWidth = 40
            .Cells(ROW_SHT_NAME_TITLE, CLM_SHT_NAME).HorizontalAlignment = xlCenter
            .Range( _
                .Cells(ROW_SHT_NAME_TITLE, CLM_SHT_NAME), _
                .Cells(.Rows.Count, CLM_SHT_NAME).End(xlUp) _
            ).Borders.LineStyle = True
            .Rows(ROW_SHT_NAME_TITLE + 1).Select
            ActiveWindow.FreezePanes = True
            .Rows(ROW_SHT_NAME_TITLE).Select
            Selection.AutoFilter
            .Cells(1, 1).Select
        End With
        
        Application.ScreenUpdating = True
    End With
End Sub

' =============================================================================
' = 概要：シートを並び替える。
' =       シート並べ替え作業用シートに記載の通り、シートを並び替える。
' =       必ずシート並べ替え作業用シートから呼び出すこと！
' =============================================================================
Public Sub SortSheetPost()
    Dim asShtName() As String
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lArrIdx As Long
    Dim lRowIdx As Long
    
    With ActiveWorkbook
        'シート名取得
        lStrtRow = ROW_SHT_NAME_STRT
        lLastRow = .Sheets(WORK_SHEET_NAME).Cells(.Sheets(WORK_SHEET_NAME).Rows.Count, CLM_SHT_NAME).End(xlUp).Row
        ReDim Preserve asShtName(lLastRow - lStrtRow)
        lArrIdx = 0
        For lRowIdx = lStrtRow To lLastRow
            asShtName(lArrIdx) = .Sheets(WORK_SHEET_NAME).Cells(lRowIdx, CLM_SHT_NAME).Value
            lArrIdx = lArrIdx + 1
        Next lRowIdx
        
        'シート数比較
        If UBound(asShtName) + 1 = .Sheets.Count - 1 Then
            'Do Nothing
        Else
            MsgBox "シート数が一致しません！"
            MsgBox "処理を中断します。"
            End
        End If
        
        Application.ScreenUpdating = False
        
        'シート並べ替え
        For lArrIdx = 0 To UBound(asShtName)
            .Sheets(asShtName(lArrIdx)).Move Before:=Sheets(lArrIdx + 1)
        Next lArrIdx
        
        '作業用シートアクティベート
        .Sheets(WORK_SHEET_NAME).Activate
        
        '作業用シート削除は暫定無効
'        '作業用シート削除
'        Application.DisplayAlerts = False
'        .Sheets(WORK_SHEET_NAME).Delete
'        Application.DisplayAlerts = True
        
        Application.ScreenUpdating = True
    End With
    
    MsgBox "並べ替え完了！"
End Sub

' =============================================================================
' = 概要：選択した範囲のハイパーリンクを一括で開く
' =============================================================================
Public Sub ハイパーリンク一括オープン()
    Dim Rng As Range
    
    If TypeName(Selection) = "Range" Then
        For Each Rng In Selection
            If Rng.Hyperlinks.Count > 0 Then Rng.Hyperlinks(1).Follow
        Next
    Else
        MsgBox "セル範囲が選択されていません。", vbExclamation
    End If
End Sub

' =============================================================================
' = 概要：フォント色を「赤」⇔「自動」でトグルする
' =============================================================================
Public Sub フォント色をトグル()
    Const COLOR_R As Long = 255
    Const COLOR_G As Long = 0
    Const COLOR_B As Long = 0
    If Selection(1).Font.Color = RGB(COLOR_R, COLOR_G, COLOR_B) Then
        Selection.Font.ColorIndex = xlAutomatic
    Else
        Selection.Font.Color = RGB(COLOR_R, COLOR_G, COLOR_B)
    End If
End Sub

' =============================================================================
' = 概要：背景色を「黄」⇔「背景色なし」でトグルする
' =============================================================================
Public Sub 背景色をトグル()
    Const COLOR_R As Long = 255
    Const COLOR_G As Long = 255
    Const COLOR_B As Long = 0
    If Selection(1).Interior.Color = RGB(COLOR_R, COLOR_G, COLOR_B) Then
        Selection.Interior.ColorIndex = 0
    Else
        Selection.Interior.Color = RGB(COLOR_R, COLOR_G, COLOR_B)
    End If
End Sub

' =============================================================================
' = 概要：オートフィルを実行する。
' =       指定した方向に応じて選択範囲を広げてオートフィルを実行する。
' =============================================================================
Public Sub オートフィル実行( _
    ByVal sDirection As String _
)
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    Dim lErrorNo As Long
    lErrorNo = 0
    
    Dim rSrc As Range
    Set rSrc = Selection
    Dim lSrcRow As Long
    Dim lSrcClm As Long
    lSrcRow = ActiveCell.Row
    lSrcClm = ActiveCell.Column
    
    '選択範囲拡大
    If lErrorNo = 0 Then
        Select Case sDirection
            Case "Right": Range(Selection, Selection.Offset(0, 1)).Select
            Case "Left": Range(Selection, Selection.Offset(0, -1)).Select
            Case "Down": Range(Selection, Selection.Offset(1, 0)).Select
            Case "Up": Range(Selection, Selection.Offset(-1, 0)).Select
            Case Else: Debug.Assert 1
        End Select
        If Err.Number = 0 Then
            'Do Nothing
        Else
            lErrorNo = 1
        End If
    Else
        'Do Nothing
    End If
    
    'オートフィル
    If lErrorNo = 0 Then
        rSrc.AutoFill Destination:=Selection
        If Err.Number = 0 Then
            'Do Nothing
        Else
            lErrorNo = 2
        End If
    Else
        'Do Nothing
    End If
    
    '画面スクロール
    If lErrorNo = 0 Then
        Select Case sDirection
            Case "Right": Selection((lSrcRow - Selection(1).Row + 1), Selection.Columns.Count).Activate
            Case "Left": Selection((lSrcRow - Selection(1).Row + 1), 1).Activate
            Case "Down": Selection(Selection.Rows.Count, (lSrcClm - Selection(1).Column + 1)).Activate
            Case "Up": Selection(1, (lSrcClm - Selection(1).Column + 1)).Activate
'            Case "Right": Cells(lSrcRow, Selection.Columns.Count).Activate
'            Case "Left": Cells(lSrcRow, 1).Activate
'            Case "Down": Cells(Selection.Rows.Count, lSrcClm).Activate
'            Case "Up": Cells(1, lSrcClm).Activate
            Case Else: Debug.Assert 1
        End Select
        If Err.Number = 0 Then
            'Do Nothing
        Else
            lErrorNo = 3
        End If
    Else
        'Do Nothing
    End If
    
    Select Case lErrorNo
        Case 0: 'Do Nothing
        Case 1: Debug.Print "【オートフィル展開<" & sDirection & ">】移動時エラー No." & Err.Number & " : " & Err.Description
        Case 2: Debug.Print "【オートフィル展開<" & sDirection & ">】オートフィル時エラー No." & Err.Number & " : " & Err.Description
        Case 3: Debug.Print "【オートフィル展開<" & sDirection & ">】スクロール時エラー No." & Err.Number & " : " & Err.Description
        Case Else: Debug.Assert 1
    End Select
    
    On Error GoTo 0
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
End Sub

' *****************************************************************************
' * 内部関数定義
' *****************************************************************************
Private Function NumConvStr2Lng( _
    ByVal sNum As String _
) As Long
    NumConvStr2Lng = Asc(sNum) + 30913
End Function

Private Function NumConvLng2Str( _
    ByVal lNum As Long _
) As String
    NumConvLng2Str = Chr(lNum - 30913)
End Function

Private Function TreeGroupSub( _
    ByRef shTrgtSht As Worksheet, _
    ByVal lGrpStrtRow As Long, _
    ByVal lGrpLastRow As Long, _
    ByVal lGrpStrtClm As Long, _
    ByVal lGrpLastClm As Long _
)
    Dim lCurRow As Long
    Dim lTrgtClm As Long
    Dim lAddRow As Long
    Dim lSubGrpStrtRow As Long
    Dim lSubGrpLastRow As Long
    Dim lSubGrpChkRow As Long
    
    Debug.Assert lGrpLastRow >= lGrpStrtRow
    Debug.Assert lGrpLastClm >= lGrpStrtClm
    
    If lGrpStrtClm >= lGrpLastClm Then
        'Do Nothing
    Else
        lCurRow = lGrpStrtRow
        lTrgtClm = lGrpStrtClm
        Do While lCurRow < lGrpLastRow
            If IsGroupParent(shTrgtSht, lCurRow, lTrgtClm) = True Then
                '=== サブグループ範囲判定 ===
                lSubGrpStrtRow = lCurRow + 1
                lSubGrpChkRow = lSubGrpStrtRow + 1
                Do While shTrgtSht.Cells(lSubGrpChkRow, lTrgtClm).Value = "" And _
                         lSubGrpChkRow <= lGrpLastRow
                    lSubGrpChkRow = lSubGrpChkRow + 1
                Loop
                lSubGrpLastRow = lSubGrpChkRow - 1
                '=== サブグループのグループ化 ===
                shTrgtSht.Range( _
                    shTrgtSht.Rows(lSubGrpStrtRow), _
                    shTrgtSht.Rows(lSubGrpLastRow) _
                ).Group
                '=== 再帰呼び出し ===
                Call TreeGroupSub( _
                    shTrgtSht, _
                    lSubGrpStrtRow, _
                    lSubGrpLastRow, _
                    lTrgtClm + 1, _
                    lGrpLastClm _
                )
                lAddRow = lSubGrpLastRow - lSubGrpStrtRow + 1
            Else
                lAddRow = 1
            End If
            lCurRow = lCurRow + lAddRow
        Loop
    End If
End Function

' 指定したセルの直下セルが空白で、右下セルが空白でない場合、
' グループの親であると判断する。
Private Function IsGroupParent( _
    ByRef shTrgtSht As Worksheet, _
    ByVal lRow As Long, _
    ByVal lClm As Long _
) As Boolean
    Dim bRetVal As Boolean
    Dim sBtmCell As String
    Dim sBtmRightCell As String
    
    sBtmCell = ActiveSheet.Cells(lRow + 1, lClm + 0).Value
    sBtmRightCell = ActiveSheet.Cells(lRow + 1, lClm + 1).Value
    
    If sBtmCell = "" And sBtmRightCell <> "" Then     'グループの親
        bRetVal = True
    ElseIf sBtmCell <> "" And sBtmRightCell = "" Then 'グループの親でない
        bRetVal = False
    Else                                              'それ以外
        Debug.Assert 0 'ありえない
    End If
    
    IsGroupParent = bRetVal
End Function

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
Private Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "フォルダを選択してください（空欄の場合は親フォルダが選択されます）"
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
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FolderExists(sSelectedPath) Then
            ShowFolderSelectDialog = sSelectedPath
        Else
            ShowFolderSelectDialog = ""
        End If
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

'コマンドを実行
Private Function ExecDosCmd( _
    ByVal sCommand As String _
) As String
    Dim oExeResult As Object
    Dim sStrOut As String
    Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c " & sCommand)
    Do While Not (oExeResult.StdOut.AtEndOfStream)
      sStrOut = sStrOut & vbNewLine & oExeResult.StdOut.ReadLine
    Loop
    ExecDosCmd = sStrOut
    Set oExeResult = Nothing
End Function
    Private Sub Test_ExecDosCmd()
        Dim sBuf As String
        sBuf = sBuf & vbNewLine & ExecDosCmd("copy C:\Users\draem_000\Desktop\test.txt C:\Users\draem_000\Desktop\test2.txt")
        MsgBox sBuf
    End Sub

'ショートカットキーを追加
Private Function AddUserDefShortcut( _
    ByVal sKey As String, _
    ByVal sMacroName As String _
)
    If Sgn(gtShortcutKey) = 0 Then
        ReDim Preserve gtShortcutKey(0)
    Else
        ReDim Preserve gtShortcutKey(UBound(gtShortcutKey) + 1)
    End If
    gtShortcutKey(UBound(gtShortcutKey)).sKey = sKey
    gtShortcutKey(UBound(gtShortcutKey)).sMacroName = sMacroName
End Function

'ショートカットキーを有効化
Private Function EnableUserDefShortcut()
    Dim lIdx As Long
    For lIdx = LBound(gtShortcutKey) To UBound(gtShortcutKey)
        Application.OnKey gtShortcutKey(lIdx).sKey, gtShortcutKey(lIdx).sMacroName
    Next lIdx
End Function

'ショートカットキーを無効化
Private Function DisableUserDefShortcut()
    Dim lIdx As Long
    For lIdx = LBound(gtShortcutKey) To UBound(gtShortcutKey)
        Application.OnKey gtShortcutKey(lIdx).sKey
    Next lIdx
End Function



