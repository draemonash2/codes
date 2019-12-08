Attribute VB_Name = "Macros"
Option Explicit

' user define macros v2.7

' =============================================================================
' =  <<マクロ一覧>>
' =    選択範囲内で中央                             選択セルに対して「選択範囲内で中央」を実行する
' =
' =    ダブルクォートを除いてコピー                 ダブルクオーテーションなしでセルをコピーする
' =    ダブルクォートを除いて追加コピー             ダブルクオーテーションなしでセルを追加コピーする
' =
' =    選択範囲をファイルエクスポート               選択範囲をファイルとしてエクスポートする。
' =    選択範囲をまとめてコマンド実行               選択範囲内のコマンドをまとめて実行する。
' =    選択範囲をそれぞれコマンド実行               選択範囲内のコマンドをそれぞれ実行する。
' =    選択範囲内の検索文字色を変更                 選択範囲内の検索文字色を変更する
' =
' =    全シート名をコピー                           ブック内のシート名を全てコピーする
' =    シート表示非表示を切り替え                   シート表示/非表示を切り替える
' =    シート並べ替え作業用シートを作成             シート並べ替え作業用シート作成
' =
' =    セル内の丸数字をデクリメント                 ②～⑮を指定して、指定番号以降をインクリメントする
' =    セル内の丸数字をインクリメント               ①～⑭を指定して、指定番号以降をデクリメントする
' =
' =    ツリーをグループ化                           ツリーグループ化する
' =    ハイパーリンク一括オープン                   選択した範囲のハイパーリンクを一括で開く
' =
' =    フォント色をトグル                           フォント色を「赤」⇔「自動」でトグルする
' =    背景色をトグル                               背景色を「黄」⇔「背景色なし」でトグルする
' =
' =    オートフィル実行                             オートフィルを実行する
' =    アクティブセルコメントのみ表示               アクティブセルコメントのみ表示する
' =    アクティブセルコメントのみ表示および移動     アクティブセルコメントのみ表示し、移動する
' =    ハイパーリンクで飛ぶ                         アクティブセルからハイパーリンク先に飛ぶ
' =    MEMOシートへジャンプ                         アクティブブックのMEMOシートへ移動する
' =
' =    Excel方眼紙                                  Excel方眼紙
' =    EpTreeの関数ツリーをExcelで取り込む          EpTreeの関数ツリーをExcelで取り込む
' =
' =    自動列幅調整                                 列幅を自動調整する
' =    自動行幅調整                                 行幅を自動調整する
' =
' =    最前面へ移動                                 最前面へ移動する
' =    最背面へ移動                                 最背面へ移動する
' =============================================================================

'******************************************************************************
'* インクルード
'******************************************************************************
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'▼▼▼Mng_Clipboard.bas/SetToClipboard()▼▼▼
'Win32API宣言
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hData As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlag As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'本来はＣ言語用の文字列コピーだが、２つ目の引数をStringとしているので変換が行われた上でコピーされる。
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
'▲▲▲Mng_Clipboard.bas/SetToClipboard()▲▲▲

'******************************************************************************
'* 設定値
'******************************************************************************
'=== セル内の丸数字をデクリメント()/セル内の丸数字をインクリメント() ===
Const NUM_MAX = 15
Const NUM_MIN = 1

'=== アクティブセルコメントのみ表示および移動() ===
Const SETTING_KEY_CMNT_VSBL_ENB As String = "CMNT_VSBL_ENB"
Const SHTCUTKEY_KEYWORD_PREFIX As String = "SHTCUTKEY"

'=== シート並べ替え作業用シートを作成() ===
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

' ==================================================================
' = 概要    ショートカットキー設定を更新する
' = 引数    sOperate    String  [in] 操作種別(Add/Update/Delete)
' = 戻値    なし
' = 覚書    なし
' = 依存    Macros.bas/UpdateShtcutSetting()
' =         AddInSetting.cls/SearchWithKey()
' = 所属    Macros.bas
' ==================================================================
Private Function UpdateShortcutKeySettings( _
    ByVal sOperate As String _
)
    ' <<ショートカットキー追加方法>>
    '   (1) 以下の追加先に「UpdateShtcutSetting()」呼び出しを追加する。
    '       第一引数にはショートカットキー、第二引数にマクロ名を指定する。
    '       ショートカットキーは Ctrl や Shift などと組み合わせて指定できる。
    '         Shift：+、Ctrl ：^、Alt  ：%
    '       詳細は以下 URL 参照。
    '         https://msdn.microsoft.com/ja-jp/library/office/ff197461.aspx
    '   (2) マクロ「ユーザー定義ショートカットキーを設定()」を実行する。
    '
    ' <<ショートカットキー解除方法>>
    '   (1) マクロ「ユーザー定義ショートカットキーを解除()」を実行する。
    
    '▼▼▼ 追加先 ▼▼▼
    Call UpdateShtcutSetting("", "選択範囲内で中央", sOperate)
    
    Call UpdateShtcutSetting("^+c", "ダブルクォートを除いてコピー", sOperate)
    Call UpdateShtcutSetting("^%c", "ダブルクォートを除いて追加コピー", sOperate)
    Call UpdateShtcutSetting("", "選択範囲をファイルエクスポート", sOperate)
    Call UpdateShtcutSetting("", "選択範囲をそれぞれコマンド実行", sOperate)
    Call UpdateShtcutSetting("", "選択範囲をまとめてコマンド実行", sOperate)
    
    Call UpdateShtcutSetting("", "全シート名をコピー", sOperate)
    Call UpdateShtcutSetting("", "シート表示非表示を切り替え", sOperate)
    Call UpdateShtcutSetting("", "シート並べ替え作業用シートを作成", sOperate)
    
    Call UpdateShtcutSetting("", "セル内の丸数字をデクリメント", sOperate)
    Call UpdateShtcutSetting("", "セル内の丸数字をインクリメント", sOperate)
    
    Call UpdateShtcutSetting("", "ツリーをグループ化", sOperate)
    Call UpdateShtcutSetting("", "ハイパーリンク一括オープン", sOperate)
    
    Call UpdateShtcutSetting("", "フォント色をトグル", sOperate)
    Call UpdateShtcutSetting("", "背景色をトグル", sOperate)
    
    Call UpdateShtcutSetting("", "Excel方眼紙", sOperate)
    Call UpdateShtcutSetting("", "EpTreeの関数ツリーをExcelで取り込む", sOperate)
    
    Call UpdateShtcutSetting("%^+{DOWN}", "'オートフィル実行(""Down"")'", sOperate)
    Call UpdateShtcutSetting("%^+{UP}", "'オートフィル実行(""Up"")'", sOperate)
    Call UpdateShtcutSetting("%^+{RIGHT}", "'オートフィル実行(""Right"")'", sOperate)
    Call UpdateShtcutSetting("%^+{LEFT}", "'オートフィル実行(""Left"")'", sOperate)
    
    Call UpdateShtcutSetting("%{F9}", "アクティブセルコメントのみ表示および移動_モード切替", sOperate)
    
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim sValue As String
    Dim bIsRet As Boolean
    bIsRet = clSetting.SearchWithKey(SETTING_KEY_CMNT_VSBL_ENB, sValue)
    If bIsRet = True Then
        If sValue = "True" Then
            Call UpdateShtcutSetting("{DOWN}", "'アクティブセルコメントのみ表示および移動(""Down"")'", sOperate)
            Call UpdateShtcutSetting("{UP}", "'アクティブセルコメントのみ表示および移動(""Up"")'", sOperate)
            Call UpdateShtcutSetting("{RIGHT}", "'アクティブセルコメントのみ表示および移動(""Right"")'", sOperate)
            Call UpdateShtcutSetting("{LEFT}", "'アクティブセルコメントのみ表示および移動(""Left"")'", sOperate)
        ElseIf sValue = "False" Then
            Call UpdateShtcutSetting("", "'アクティブセルコメントのみ表示および移動(""Down"")'", sOperate)
            Call UpdateShtcutSetting("", "'アクティブセルコメントのみ表示および移動(""Up"")'", sOperate)
            Call UpdateShtcutSetting("", "'アクティブセルコメントのみ表示および移動(""Right"")'", sOperate)
            Call UpdateShtcutSetting("", "'アクティブセルコメントのみ表示および移動(""Left"")'", sOperate)
        Else
            Debug.Assert False
        End If
    Else
        Call UpdateShtcutSetting("", "'アクティブセルコメントのみ表示および移動(""Down"")'", sOperate)
        Call UpdateShtcutSetting("", "'アクティブセルコメントのみ表示および移動(""Up"")'", sOperate)
        Call UpdateShtcutSetting("", "'アクティブセルコメントのみ表示および移動(""Right"")'", sOperate)
        Call UpdateShtcutSetting("", "'アクティブセルコメントのみ表示および移動(""Left"")'", sOperate)
    End If
    
    Call UpdateShtcutSetting("^+j", "ハイパーリンクで飛ぶ", sOperate)
    Call UpdateShtcutSetting("^%{HOME}", "MEMOシートへジャンプ", sOperate)
    
    Call UpdateShtcutSetting("", "Excel方眼紙", sOperate)
    
    Call UpdateShtcutSetting("", "自動列幅調整", sOperate)
    Call UpdateShtcutSetting("", "自動行幅調整", sOperate)
    
    Call UpdateShtcutSetting("^+f", "最前面へ移動", sOperate)
    Call UpdateShtcutSetting("^+b", "最背面へ移動", sOperate)
    '▲▲▲ 追加先 ▲▲▲
End Function

' *****************************************************************************
' * 外部公開用マクロ
' *****************************************************************************
' =============================================================================
' = 概要    選択セルに対して「選択範囲内で中央」を実行する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 選択範囲内で中央()
    Selection.HorizontalAlignment = xlCenterAcrossSelection
End Sub

' =============================================================================
' = 概要    ①～⑭を指定して、指定番号以降をデクリメントする
' = 覚書    なし
' = 依存    Mng_String.bas/NumConvStr2Lng()
' =         Mng_String.bas/NumConvLng2Str()
' = 所属    Macros.bas
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
' = 概要    ②～⑮を指定して、指定番号以降をインクリメントする
' = 覚書    なし
' = 依存    Mng_String.bas/NumConvStr2Lng()
' =         Mng_String.bas/NumConvLng2Str()
' = 所属    Macros.bas
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
' = 概要    ブック内のシート名を全てコピーする
' = 覚書    本マクロがエラーとなる場合、以下のいずれかを実施すること。
' =           ・ツール->参照設定 にて「Microsoft Forms 2.0 Object Library」を選択
' =           ・ツール->参照設定 内の「参照」にて system32 内の「FM20.DLL」を選択
' = 依存    なし
' = 所属    Macros.bas
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
' = 概要    シート表示/非表示を切り替える
' = 覚書    なし
' = 依存    SheetVisibleSetting.cls/SheetVisibleSetting()
' = 所属    Macros.bas
' =============================================================================
Public Sub シート表示非表示を切り替え()
    SheetVisibleSetting.Show
End Sub

' =============================================================================
' = 概要    ダブルクオーテーションなしでセルコピーする
' =         非表示セルは無視する。複数範囲は未対応。
' = 覚書    なし
' = 依存    Mng_Array.bas/ConvRange2Array()
' =         Mng_Clipboard.bas/SetToClipboard()
' = 所属    Macros.bas
' =============================================================================
Public Sub ダブルクォートを除いてコピー()
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
    Call SetToClipboard(sBuf)
    
    'フィードバック
    Application.StatusBar = "■■■■■■■■ コピー完了！ ■■■■■■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' ==================================================================
' = 概要    選択範囲をクリップボードへ追加コピー
' =         ダブルクオーテーションなしでセルコピーする
' =         非表示セルは無視する。複数範囲は未対応。
' = 覚書    なし
' = 依存    Mng_Array.bas/ConvRange2Array()
' =         Mng_Clipboard.bas/SetToClipboard()
' = 所属    Macros.bas
' ==================================================================
Public Sub ダブルクォートを除いて追加コピー()
    Dim bIsInvisibleCellIgnore As Boolean
    Dim sDelimiter As String
    bIsInvisibleCellIgnore = True
    sDelimiter = Chr(9)
    
    '*** 既存テキスト取得 ***
    Dim sOrgText As String
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        sOrgText = .GetText
    End With
    
    '*** 追加テキスト取得 ***
    Dim asLine() As String
    Call ConvRange2Array( _
        Selection, _
        asLine, _
        bIsInvisibleCellIgnore, _
        sDelimiter _
    )
    
    Dim sNewText As String
    sNewText = ""
    Dim lLineIdx As Long
    For lLineIdx = LBound(asLine) To UBound(asLine)
        If lLineIdx = LBound(asLine) Then
            sNewText = asLine(lLineIdx)
        Else
            sNewText = sNewText & vbNewLine & asLine(lLineIdx)
        End If
    Next lLineIdx
    
    '*** クリップボード設定 ***
    Dim sOutText As String
    sOutText = sOrgText & vbNewLine & sNewText
    Call SetToClipboard(sOutText)

    '*** フィードバック ***
    Application.StatusBar = "■■■■■■■■ 追加コピー完了！ ■■■■■■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' =============================================================================
' = 概要    選択範囲をファイルとしてエクスポートする。
' =         隣り合った列のセルにはタブ文字を挿入して出力する。
' = 覚書    なし
' = 依存    Mng_FileSys.bas/ShowFolderSelectDialog()
' =         Mng_Array.bas/ConvRange2Array()
' = 所属    Macros.bas
' =============================================================================
Public Sub 選択範囲をファイルエクスポート()
    Const TEMP_FILE_NAME As String = "ExportCellRange.tmp"
    Const FILE_EXTENTION As String = "csv"
    Const DELIMITER As String = ","
    
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
    sTmpPath = "C:\Users\" & CreateObject("WScript.Network").UserName & "\AppData\Local\Temp\" & TEMP_FILE_NAME
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
    sOutputFilePath = sOutputDirPath & "\" & sOutputFileName & "." & FILE_EXTENTION
    
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
    'Range型からString()型へ変換
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIsInvisibleCellIgnore, _
                DELIMITER _
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
' = 概要    選択範囲内のコマンドをバッチファイルに書き出してまとめて実行する。
' =         単一列選択時のみ有効。
' = 覚書    なし
' = 依存    Mng_Array.bas/ConvRange2Array()
' =         Mng_FileSys.bas/OutputTxtFile()
' =         Mng_SysCmd.bas/ExecDosCmd()
' = 所属    Macros.bas
' =============================================================================
Public Sub 選択範囲をまとめてコマンド実行()
    Const BAT_FILE_NAME As String = "command.bat"
    
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
    
    Dim sBatFileDirPath As String
    Dim sBatFilePath As String
    sBatFileDirPath = "C:\Users\" & CreateObject("WScript.Network").UserName & "\AppData\Local\Temp"
    sBatFilePath = sBatFileDirPath & "\" & BAT_FILE_NAME
    
    Call OutputTxtFile(sBatFilePath, asRange)
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\redirect.log"
    
    '*** コマンド実行 ***
    Open sOutputFilePath For Append As #1
    Print #1, "****************************************************"
    Print #1, Now()
    Print #1, "****************************************************"
    Print #1, ExecDosCmd(sBatFilePath)
    Print #1, ""
    Close #1
    
    '*** バッチファイル削除 ***
    Kill sBatFilePath
    
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
' = 概要    選択範囲内のコマンドをそれぞれ実行する。
' =         単一列選択時のみ有効。
' = 覚書    なし
' = 依存    Mng_Array.bas/ConvRange2Array()
' =         Mng_SysCmd.bas/ExecDosCmd()
' = 所属    Macros.bas
' =============================================================================
Public Sub 選択範囲をそれぞれコマンド実行()
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
' = 概要    選択範囲内の検索文字色を変更する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 選択範囲内の検索文字色を変更()
    Const sMACRO_TITLE As String = "選択範囲内の検索文字色を変更"
    
    '▼▼▼色設定▼▼▼
    Const sCOLOR_TYPE As String = "0:赤、1:水、2:緑、3:紫、4:橙、5:黄、6:白、7:黒"
    Const lCOLOR_NUM As Long = 8
    Dim vCOLOR_INFO() As Variant
    vCOLOR_INFO = _
        Array( _
            Array(255, 0, 0), _
            Array(75, 172, 198), _
            Array(118, 147, 60), _
            Array(112, 48, 160), _
            Array(247, 150, 70), _
            Array(255, 192, 0), _
            Array(255, 255, 255), _
            Array(0, 0, 0) _
        )
    '▲▲▲色設定▲▲▲
    
    Dim sSrchStr As String
    sSrchStr = InputBox("検索文字列を入力してください", sMACRO_TITLE)
    
    Dim lColorIndex As Long
    lColorIndex = InputBox( _
        "文字色を選択してください" & vbNewLine & _
        "  " & sCOLOR_TYPE & vbNewLine _
        , _
        sMACRO_TITLE, _
        1 _
    )
    
    If lColorIndex < lCOLOR_NUM Then
        Dim oCell As Range
        For Each oCell In Selection
            Dim sTrgtStr As String
            sTrgtStr = oCell.Value
            Dim lStartIdx As Long
            lStartIdx = 1
            Do
                Dim lIdx As Long
                lIdx = InStr(lStartIdx, sTrgtStr, sSrchStr)
                If lIdx = 0 Then
                    Exit Do
                Else
                    lStartIdx = lIdx + Len(sSrchStr)
                    oCell.Characters(Start:=lIdx, Length:=Len(sSrchStr)).Font.Color = _
                        RGB( _
                            vCOLOR_INFO(lColorIndex)(0), _
                            vCOLOR_INFO(lColorIndex)(1), _
                            vCOLOR_INFO(lColorIndex)(2) _
                        )
                End If
            Loop While 1
        Next
        MsgBox "完了！", vbOKOnly, sMACRO_TITLE
    Else
        MsgBox "文字色は指定の範囲内で選択してください。" & vbNewLine & sCOLOR_TYPE, vbOKOnly, sMACRO_TITLE
    End If
End Sub

' =============================================================================
' = 概要    行をツリー構造にしてグループ化
' =         Usage：ツリーグループ化したい範囲を選択し、マクロ「ツリーをグループ化」を実行する
' = 覚書    なし
' = 依存    Macros.bas/TreeGroupSub()
' = 所属    Macros.bas
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
' = 概要    シートを並び替える。
' =         本処理を実行すると、シート並べ替え作業用シートを作成する。
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
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
            Set shWorkSht = .Sheets.Add(after:=.Sheets(.Sheets.Count))
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
' = 概要    シートを並び替える。
' =         シート並べ替え作業用シートに記載の通り、シートを並び替える。
' =         必ずシート並べ替え作業用シートから呼び出すこと！
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
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
' = 概要    選択した範囲のハイパーリンクを一括で開く
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
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
' = 概要    フォント色を「赤」⇔「自動」でトグルする
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
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
' = 概要    背景色を「黄」⇔「背景色なし」でトグルする
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
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
' = 概要    オートフィルを実行する。
' =         指定した方向に応じて選択範囲を広げてオートフィルを実行する。
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
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

' =============================================================================
' = 概要    アクティブセルコメントのみ表示し、移動する。
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub アクティブセルコメントのみ表示および移動( _
    ByVal sDirection As String _
)
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    'アクティブセルコメント表示
    Dim cmComment As Comment
    For Each cmComment In ActiveSheet.Comments
        cmComment.Visible = False
    Next cmComment
    
    'セル移動
    Select Case sDirection
        Case "Right": ActiveCell.Offset(0, 1).Activate
        Case "Left": ActiveCell.Offset(0, -1).Activate
        Case "Down": ActiveCell.Offset(1, 0).Activate
        Case "Up": ActiveCell.Offset(-1, 0).Activate
        Case Else: Debug.Assert 1
    End Select
    
    'アクティブセルコメント表示
    ActiveCell.Comment.Visible = True
    
    On Error GoTo 0
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = 概要    アクティブセルコメントのみ表示する。
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub アクティブセルコメントのみ表示()
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    'アクティブセルコメント表示
    Dim cmComment As Comment
    For Each cmComment In ActiveSheet.Comments
        cmComment.Visible = False
    Next cmComment
    
    'アクティブセルコメント表示
    ActiveCell.Comment.Visible = True
    
    On Error GoTo 0
    
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = 概要    アクティブセルのコメント表示の有効/無効を切り替える
' = 覚書    なし
' = 依存    AddInSetting.cls/SearchWithKey()
' =         AddInSetting.cls/Update()
' =         AddInSetting.cls/Add()
' =         Macros.bas/DisableShortcutKeys()
' =         Macros.bas/UpdateShortcutKeySettings()
' =         Macros.bas/EnableShortcutKeys()
' = 所属    Macros.bas
' =============================================================================
Public Sub アクティブセルコメントのみ表示および移動_モード切替()
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim bRet As Boolean
    Dim sValue As String
    bRet = clSetting.SearchWithKey(SETTING_KEY_CMNT_VSBL_ENB, sValue)
    If bRet = True Then
        If sValue = "True" Then
            MsgBox "アクティブセルコメントのみ表示および移動を無効にします"
            Call DisableShortcutKeys
            Call clSetting.Update(SETTING_KEY_CMNT_VSBL_ENB, "False")
            Call UpdateShortcutKeySettings("Update")
            Call EnableShortcutKeys
        Else
            MsgBox "アクティブセルコメントのみ表示および移動を有効にします"
            Call DisableShortcutKeys
            Call clSetting.Update(SETTING_KEY_CMNT_VSBL_ENB, "True")
            Call UpdateShortcutKeySettings("Update")
            Call EnableShortcutKeys
        End If
        Debug.Assert bRet
    Else
        MsgBox "アクティブセルコメントのみ表示および移動を無効にします"
        Call DisableShortcutKeys
        Call clSetting.Add(SETTING_KEY_CMNT_VSBL_ENB, "False")
        Call UpdateShortcutKeySettings("Update")
        Call EnableShortcutKeys
    End If
End Sub

' =============================================================================
' = 概要    アクティブセルからハイパーリンク先に飛ぶ
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub ハイパーリンクで飛ぶ()
    On Error Resume Next
    ActiveCell.Hyperlinks(1).Follow NewWindow:=True
    If Err.Number = 0 Then
        'Do Nothing
    Else
        Debug.Print "[" & Now & "] Error " & _
                    "[Macro] ハイパーリンクで飛ぶ " & _
                    "[Error No." & Err.Number & "] " & Err.Description
    End If
    On Error GoTo 0
End Sub

' =============================================================================
' = 概要    アクティブブックのMEMOシートへ移動する
' = 覚書    なし
' = 依存　　Macros.bas/JumpToTrgtSheet()
' = 所属    Macros.bas
' =============================================================================
Public Sub MEMOシートへジャンプ()
    Const TRGT_SHEET_NAME As String = "MEMO"
    Call JumpToTrgtSheet(TRGT_SHEET_NAME)
End Sub

' =============================================================================
' = 概要    Excel方眼紙
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub Excel方眼紙()
    ActiveSheet.Cells.Select
    With Selection
        .ColumnWidth = 1.22
        .RowHeight = 10.8
        .Font.Size = 9
        .Font.Name = "ＭＳ ゴシック"
    End With
    ActiveSheet.Cells(1, 1).Select
End Sub

' =============================================================================
' = 概要    列幅、行幅を自動調整する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 自動列幅調整()
    Selection.EntireColumn.AutoFit
End Sub
Public Sub 自動行幅調整()
    Selection.EntireRow.AutoFit
End Sub

' =============================================================================
' = 概要    最前面、最背面へ移動する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 最前面へ移動()
    Selection.ShapeRange.ZOrder msoBringToFront
End Sub
Public Sub 最背面へ移動()
    Selection.ShapeRange.ZOrder msoSendToBack
End Sub

' =============================================================================
' = 概要    EpTreeの関数ツリーをExcelで取り込む
' = 覚書    なし
' = 依存    Mng_FileSys.bas/ShowFilesSelectDialog()
' =         Mng_Collection.bas/ReadTxtFileToCollection()
' =         Mng_String.bas/ExecRegExp()
' =         Mng_ExcelOpe.bas/CreateNewWorksheet()
' = 所属    Macros.bas
' =============================================================================
Public Sub EpTreeの関数ツリーをExcelで取り込む()
    Const STRT_ROW As Long = 2
    Const STRT_CLM As Long = 2
    Const SHEET_NAME As String = "関数ツリー"
    Const MAX_FUNC_LEVEL_INI As Long = 10
    Const CLM_WIDTH As Long = 2
    Dim lRowIdx As Long
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lStrtClm As Long
    Dim lLastClm As Long
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'ファイルパス入力
    Dim sTrgtFilePath As String
    Dim asSelectedFiles() As String
    Call ShowFilesSelectDialog(asSelectedFiles)
    sTrgtFilePath = asSelectedFiles(0)
    
    'シート追加
    Dim sSheetName As String
    Dim shTrgtSht As Worksheet
    sSheetName = CreateNewWorksheet(SHEET_NAME)
    Set shTrgtSht = ActiveWorkbook.Sheets(sSheetName)
    
    'テキストファイル読み出し
    Dim cFileContents As Collection
    Set cFileContents = New Collection
    Call ReadTxtFileToCollection(sTrgtFilePath, cFileContents)
    
    'ファイルツリー出力
    lStrtRow = STRT_ROW
    lStrtClm = STRT_CLM
    lRowIdx = lStrtRow
    
    With shTrgtSht
        .Cells(lRowIdx, lStrtClm + 0).Value = "ファイルパス"
        .Cells(lRowIdx, lStrtClm + 1).Value = "行数"
        .Cells(lRowIdx, lStrtClm + 2).Value = "関数名"
        .Cells(lRowIdx, lStrtClm + 3).Value = "関数ツリー"
    End With
    lRowIdx = lRowIdx + 1
    
    Dim lMaxFuncLevel As Long
    lMaxFuncLevel = MAX_FUNC_LEVEL_INI
    Dim vFileLine As Variant
    For Each vFileLine In cFileContents
        Dim oMatchResult As Object
        Call ExecRegExp( _
            vFileLine, _
            "^(.+)? +(\d+): (  )?([│|└|├|  ]*)(\w+)(↑)?", _
            oMatchResult _
        )
        
        Dim sFilePath As String
        Dim sLineNo As String
        Dim lFuncLevel As Long
        Dim sFuncName As String
        Dim sOmission As String
        sFilePath = oMatchResult(0).SubMatches(0)
        sLineNo = oMatchResult(0).SubMatches(1)
        If sLineNo = 0 Then
            sLineNo = ""
        End If
        lFuncLevel = LenB(oMatchResult(0).SubMatches(3)) / 2
        sFuncName = oMatchResult(0).SubMatches(4)
        sOmission = String(LenB(oMatchResult(0).SubMatches(5)) / 2, "▲")
        
        With shTrgtSht
            .Cells(lRowIdx, lStrtClm + 0).Value = sFilePath
            .Cells(lRowIdx, lStrtClm + 1).Value = sLineNo
            .Cells(lRowIdx, lStrtClm + 2).Value = sFuncName
            .Cells(lRowIdx, lStrtClm + 3 + lFuncLevel).Value = sFuncName & sOmission
        End With
        If lFuncLevel > lMaxFuncLevel Then
            lMaxFuncLevel = lFuncLevel
        End If
        
        lRowIdx = lRowIdx + 1
    Next
    
    With shTrgtSht
        lLastClm = STRT_CLM + 3 + lMaxFuncLevel
        lLastRow = lRowIdx
        
        'タイトル行 中央揃え
        .Range(.Cells(lStrtRow, lStrtClm + 0), .Cells(lStrtRow, lStrtClm + 2)).HorizontalAlignment = xlCenter
        .Range(.Cells(lStrtRow, lStrtClm + 3), .Cells(lStrtRow, lLastClm)).HorizontalAlignment = xlCenterAcrossSelection
        
        '列幅調整
        .Range(.Cells(lStrtRow, lStrtClm + 0), .Cells(lLastRow, lStrtClm + 0)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 1), .Cells(lLastRow, lStrtClm + 1)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 2), .Cells(lLastRow, lStrtClm + 2)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 3), .Cells(lLastRow, lLastClm)).ColumnWidth = CLM_WIDTH
    
        'オートフィルタ
        .Range(.Cells(lStrtRow, lStrtClm), .Cells(lLastRow, lLastClm)).AutoFilter
        
        '行高さ
        .Rows(lStrtRow).RowHeight = .Rows(lStrtRow).RowHeight * 3
        
        'タイトル列固定
        ActiveWindow.FreezePanes = False
        .Rows(lStrtRow + 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1).Select
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "関数ツリー作成完了！"
End Sub

' *****************************************************************************
' * 内部用マクロ
' *****************************************************************************
' =============================================================================
' = 概要    設定項目一覧を出力
' = 覚書    なし
' = 依存    AddInSetting.cls/SearchWithIdx()
' = 所属    Macros.bas
' =============================================================================
Private Sub OutputSettingList()
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim lSettingNum As Long
    lSettingNum = clSetting.Count
    
    Debug.Print ""
    Debug.Print "*** 設定項目一覧出力 ***"
    If lSettingNum = 0 Then
        'Do Nothing
    Else
        Dim lSettingIdx As Long
        For lSettingIdx = 1 To lSettingNum
            Dim sSettingKey As String
            Dim sSettingValue As String
            Call clSetting.SearchWithIdx(lSettingIdx, sSettingKey, sSettingValue)
            Debug.Print sSettingKey & " = " & sSettingValue
        Next lSettingIdx
    End If
End Sub

' =============================================================================
' = 概要    ユーザー定義ショートカットキー設定を追加
' = 覚書    なし
' = 依存    Macros.bas/UpdateShortcutKeySettings()
' = 所属    Macros.bas
' =============================================================================
Public Sub ユーザー定義ショートカットキー設定を追加()
    Call UpdateShortcutKeySettings("Add")
End Sub

' =============================================================================
' = 概要    ユーザー定義ショートカットキー設定を削除
' = 覚書    なし
' = 依存    Macros.bas/UpdateShortcutKeySettings()
' = 所属    Macros.bas
' =============================================================================
Public Sub ユーザー定義ショートカットキー設定を削除()
    Call UpdateShortcutKeySettings("Delete")
End Sub

' =============================================================================
' = 概要    ユーザー定義ショートカットキー設定を更新
' = 覚書    なし
' = 依存    Macros.bas/UpdateShortcutKeySettings()
' = 所属    Macros.bas
' =============================================================================
Public Sub ユーザー定義ショートカットキー設定を更新()
    Call UpdateShortcutKeySettings("Update")
End Sub

' =============================================================================
' = 概要    ユーザー定義ショートカットキーを有効化
' = 覚書    なし
' = 依存    Macros.bas/EnableShortcutKeys()
' = 所属    Macros.bas
' =============================================================================
Public Sub ユーザー定義ショートカットキーを有効化()
    Call EnableShortcutKeys
End Sub

' =============================================================================
' = 概要    ユーザー定義ショートカットキーを無効化
' = 覚書    なし
' = 依存    Macros.bas/DisableShortcutKeys()
' = 所属    Macros.bas
' =============================================================================
Public Sub ユーザー定義ショートカットキーを無効化()
    Call DisableShortcutKeys
End Sub

' *****************************************************************************
' * 内部関数定義
' *****************************************************************************
' ==================================================================
' = 概要    数字 型変換(String→Long)
' = 引数    sNum            String  [in]  数字(String型)
' = 戻値                    Long          数字(Long型)
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Private Function NumConvStr2Lng( _
    ByVal sNum As String _
) As Long
    NumConvStr2Lng = Asc(sNum) + 30913
End Function

' ==================================================================
' = 概要    数字 型変換(Long→String)
' = 引数    lNum            Long    [in]    数字(Long型)
' = 戻値                    String          数字(String型)
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Private Function NumConvLng2Str( _
    ByVal lNum As Long _
) As String
    NumConvLng2Str = Chr(lNum - 30913)
End Function

' ==================================================================
' = 概要    ツリーをグループ化
' = 引数    shTrgtSht       Worksheet   [in,out]    ワークシート
' = 引数    lGrpStrtRow     Long        [in]        先頭行
' = 引数    lGrpLastRow     Long        [in]        末尾行
' = 引数    lGrpStrtClm     Long        [in]        先頭列
' = 引数    lGrpLastClm     Long        [in]        末尾列
' = 戻値    なし
' = 覚書    なし
' = 依存    Macros.bas/IsGroupParent()
' = 所属    Macros.bas
' ==================================================================
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

' ==================================================================
' = 概要    指定したセルの直下セルが空白で、右下セルが空白でない場合、
' =         グループの親であると判断する。
' = 引数    shTrgtSht   Worksheet   [in,out]    ワークシート
' = 引数    lRow        Long        [in]        行
' = 引数    lClm        Long        [in]        列
' = 戻値                Boolean
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
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
' = 依存    なし
' = 所属    Mng_Array.bas
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
' = 依存    なし
' = 所属    Mng_FileSys.bas
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

' ==================================================================
' = 概要    コマンドを実行
' = 引数    sCommand    String   [in]   コマンド
' = 戻値                String          標準出力
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_SysCmd.bas
' ==================================================================
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

' ==================================================================
' = 概要    ショートカットキー設定を追加/削除
' = 引数    sKey        String   [in]   キー
' = 引数    sMacroName  String   [in]   マクロ名
' = 引数    sMode       String   [in]   モード(Add/Update/Delete)
' = 戻値    なし
' = 覚書    なし
' = 依存    AddInSetting.cls/Add()
' =         AddInSetting.cls/Update()
' =         AddInSetting.cls/Delete()
' = 所属    Macros.bas
' ==================================================================
Private Function UpdateShtcutSetting( _
    ByVal sKey As String, _
    ByVal sMacroName As String, _
    ByVal sMode As String _
)
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim sSettingKey As String
    Dim sSettingValue As String
    sSettingKey = SHTCUTKEY_KEYWORD_PREFIX & "_" & sMacroName
    sSettingValue = sKey
    Select Case sMode
        Case "Add": Call clSetting.Add(sSettingKey, sSettingValue)
        Case "Update": Call clSetting.Update(sSettingKey, sSettingValue)
        Case "Delete": Call clSetting.Delete(sSettingKey)
        Case Else: Debug.Assert False
    End Select
End Function

' ==================================================================
' = 概要    ショートカットキーを有効化
' = 引数    なし
' = 戻値    なし
' = 覚書    なし
' = 依存    AddInSetting.cls/Count()
' =         AddInSetting.cls/SearchWithIdx()
' = 所属    Macros.bas
' ==================================================================
Private Function EnableShortcutKeys()
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim lNum As Long
    lNum = clSetting.Count
    If lNum = 0 Then
        'Do Nothing
    Else
        Dim lLastRow As Long
        lLastRow = lNum
        Dim lRowIdx As Long
        For lRowIdx = 1 To lLastRow
            Dim sSettingKey As String
            Dim sSettingValue As String
            Call clSetting.SearchWithIdx(lRowIdx, sSettingKey, sSettingValue)
            '*** ショートカット設定の場合 ***
            If InStr(sSettingKey, SHTCUTKEY_KEYWORD_PREFIX) Then
                Dim sShrcutMacroName As String
                Dim sShtcutKey As String
                sShrcutMacroName = Replace(sSettingKey, SHTCUTKEY_KEYWORD_PREFIX & "_", "")
                sShtcutKey = sSettingValue
                If sShtcutKey = "" Then
                    'Do Nothing
                Else
                    Application.OnKey sShtcutKey, sShrcutMacroName
                End If
            
            '*** ショートカット設定でない場合 ***
            Else
                'Do Nothing
            End If
        Next lRowIdx
    End If
End Function

' ==================================================================
' = 概要    ショートカットキーを無効化
' = 引数    なし
' = 戻値    なし
' = 覚書    なし
' = 依存    AddInSetting.cls/Count()
' =         AddInSetting.cls/SearchWithIdx()
' = 所属    Macros.bas
' ==================================================================
Private Function DisableShortcutKeys()
    Dim clSetting As AddInSetting
    Set clSetting = New AddInSetting
    Dim lNum As Long
    lNum = clSetting.Count
    If lNum = 0 Then
        'Do Nothing
    Else
        Dim lLastRow As Long
        lLastRow = lNum
        Dim lRowIdx As Long
        For lRowIdx = 1 To lLastRow
            Dim sSettingKey As String
            Dim sSettingValue As String
            Call clSetting.SearchWithIdx(lRowIdx, sSettingKey, sSettingValue)
            If InStr(sSettingKey, SHTCUTKEY_KEYWORD_PREFIX) Then
                Dim sShtcutKey As String
                sShtcutKey = sSettingValue
                If sShtcutKey = "" Then
                    'Do Nothing
                Else
                    Application.OnKey sShtcutKey
                End If
            Else
                'Do Nothing
            End If
        Next lRowIdx
    End If
End Function

' ============================================
' = 概要    配列の内容をファイルに書き込む。
' = 引数    sFilePath     String  [in]  出力するファイルパス
' =         asFileLine()  String  [in]  出力するファイルの内容
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_Array.bas
' ============================================
Private Function OutputTxtFile( _
    ByVal sFilePath As String, _
    ByRef asFileLine() As String, _
    Optional ByVal sCharSet As String = "shift_jis" _
)
    Dim oTxtObj As Object
    Dim lLineIdx As Long
    
    If Sgn(asFileLine) = 0 Then
        'Do Nothing
    Else
        Set oTxtObj = CreateObject("ADODB.Stream")
        With oTxtObj
            .Type = 2
            .Charset = sCharSet
            .Open
            
            '配列を1行ずつオブジェクトに書き込む
            For lLineIdx = 0 To UBound(asFileLine)
                .WriteText asFileLine(lLineIdx), 1
            Next lLineIdx
            
            .SaveToFile (sFilePath), 2    'オブジェクトの内容をファイルに保存
            .Close
        End With
    End If
    
    Set oTxtObj = Nothing
End Function

' ==================================================================
' = 概要    ファイル（複数）選択ダイアログを表示する
' = 引数    asSelectedFiles String()    [out] 選択されたファイルパス一覧
' = 引数    sInitPath       String      [in]  デフォルトファイルパス（省略可）
' = 引数    sFilters        String      [in]  選択時のフィルタ（省略可）(※)
' = 戻値    なし
' = 覚書    (※)ダイアログのフィルタ指定方法は以下。
' =              ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =                    ・拡張子が複数ある場合は、";"で区切る
' =                    ・ファイル種別と拡張子は"/"で区切る
' =                    ・フィルタが複数ある場合、","で区切る
' =         sFilters が省略もしくは空文字の場合、フィルタをクリアする。
' = 依存    Mng_FileSys.bas/SetDialogFilters()
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function ShowFilesSelectDialog( _
    ByRef asSelectedFiles() As String, _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFilters As String = "" _
)
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "ファイルを選択してください（複数可）"
    fdDialog.AllowMultiSelect = True
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) 'フィルタ追加
 
    'ダイアログ表示
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then 'キャンセル押下
        ReDim Preserve asSelectedFiles(0)
        asSelectedFiles(0) = ""
    Else
        Dim lSelNum As Long
        lSelNum = fdDialog.SelectedItems.Count
        ReDim Preserve asSelectedFiles(lSelNum - 1)
        Dim lSelIdx As Long
        For lSelIdx = 0 To lSelNum - 1
            Dim sSelectedPath As String
            sSelectedPath = fdDialog.SelectedItems(lSelIdx + 1)
            If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
                asSelectedFiles(lSelIdx) = sSelectedPath
            Else
                asSelectedFiles(lSelIdx) = ""
            End If
        Next lSelIdx
    End If
 
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFilesSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sFilters As String
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv"
        'sFilters = "画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
        sFilters = "全てのファイル/*.*,画像ファイル/*.gif; *.jpg; *.jpeg; *.png,テキストファイル/*.txt; *.csv"
 
        Dim asSelectedFiles() As String
        Call ShowFilesSelectDialog( _
                    asSelectedFiles, _
                    objWshShell.SpecialFolders("Desktop") & "\test.txt", _
                    sFilters _
                )
        Dim sBuf As String
        sBuf = ""
        sBuf = sBuf & vbNewLine & UBound(asSelectedFiles) + 1
        Dim lSelIdx As Long
        For lSelIdx = 0 To UBound(asSelectedFiles)
            sBuf = sBuf & vbNewLine & asSelectedFiles(lSelIdx)
        Next lSelIdx
        MsgBox sBuf
    End Sub
 
' ==================================================================
' = 概要    ShowFileSelectDialog() と ShowFilesSelectDialog() 用の関数
' =         ダイアログのフィルタを追加する。指定方法は以下。
' =           ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =               ・拡張子が複数ある場合は、";"で区切る
' =               ・ファイル種別と拡張子は"/"で区切る
' =               ・フィルタが複数ある場合、","で区切る
' =         sFilters が空文字の場合、フィルタをクリアする。
' = 引数    sFilters    String  [in]    フィルタ
' = 引数    fdDialog    String  [out]   ダイアログ
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function SetDialogFilters( _
    ByVal sFilters As String, _
    ByRef fdDialog As FileDialog _
)
    fdDialog.Filters.Clear
    If sFilters = "" Then
        'Do Nothing
    Else
        Dim vFilter As Variant
        If InStr(sFilters, ",") > 0 Then
            Dim vFilters As Variant
            vFilters = Split(sFilters, ",")
            Dim lFilterIdx As Long
            For lFilterIdx = 0 To UBound(vFilters)
                If InStr(vFilters(lFilterIdx), "/") > 0 Then
                    vFilter = Split(vFilters(lFilterIdx), "/")
                    If UBound(vFilter) = 1 Then
                        fdDialog.Filters.Add vFilter(0), vFilter(1), lFilterIdx + 1
                    Else
                        MsgBox _
                            "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                            """/"" は一つだけ指定してください" & vbNewLine & _
                            "  " & vFilters(lFilterIdx)
                        MsgBox "処理を中断します。"
                        End
                    End If
                Else
                    MsgBox _
                        "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                        "種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
                        "  " & vFilters(lFilterIdx)
                    MsgBox "処理を中断します。"
                    End
                End If
            Next lFilterIdx
        Else
            If InStr(sFilters, "/") > 0 Then
                vFilter = Split(sFilters, "/")
                If UBound(vFilter) = 1 Then
                    fdDialog.Filters.Add vFilter(0), vFilter(1), 1
                Else
                    MsgBox _
                        "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                        """/"" は一つだけ指定してください" & vbNewLine & _
                        "  " & sFilters
                    MsgBox "処理を中断します。"
                    End
                End If
            Else
                MsgBox _
                    "ファイル選択ダイアログのフィルタの指定方法が誤っています" & vbNewLine & _
                    "種別と拡張子を ""/"" で区切ってください。" & vbNewLine & _
                    "  " & sFilters
                MsgBox "処理を中断します。"
                End
            End If
        End If
    End If
End Function

' ==================================================================
' = 概要    ワークシートを新規作成
' =         重複したワークシートがある場合、_1, _2 ...と連番になる。
' =         呼び出し側には作成したワークシート名を返す。
' = 引数    sSheetName  String  [in]    シート名
' = 戻値                                シート名
' = 覚書    なし
' = 依存    Mng_ExcelOpe.bas/ExistsWorksheet()
' = 所属    Mng_ExcelOpe.bas
' ==================================================================
Private Function CreateNewWorksheet( _
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
' = 引数    sTrgtShtName    String  [in]    シート名
' = 戻値                                    存在チェック結果
' = 覚書    なし
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
' = 概要    テキストファイルの中身を配列に格納
' = 引数    sTrgtFilePath   String      [in]    ファイルパス
' = 引数    cFileContents   Collections [out]   ファイルの中身
' = 戻値    読み出し結果    Boolean             読み出し結果
' =                                                 True:ファイル存在
' =                                                 False:それ以外
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_Collection.bas
' ==================================================================
Private Function ReadTxtFileToCollection( _
    ByVal sTrgtFilePath As String, _
    ByRef cFileContents As Collection _
)
    On Error Resume Next
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(sTrgtFilePath) Then
        Dim objTxtFile As Object
        Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
        
        If Err.Number = 0 Then
            Do Until objTxtFile.AtEndOfStream
                cFileContents.Add objTxtFile.ReadLine
            Loop
            ReadTxtFileToCollection = True
        Else
            ReadTxtFileToCollection = False
        '   WScript.Echo "エラー " & Err.Description
        End If
        
        objTxtFile.Close
    Else
        ReadTxtFileToCollection = False
    End If
    On Error GoTo 0
End Function
    Private Sub Test_OpenTxtFile2Array()
        Dim cFileContents As Collection
        Set cFileContents = New Collection
        Dim sInFilePath As String
        sInFilePath = "C:\codes\vbs\_lib\Test.csv"
        Dim bRet As Boolean
        bRet = ReadTxtFileToCollection(sInFilePath, cFileContents)
    End Sub

' ==================================================================
' = 概要    正規表現検索を行う（Vbaマクロ関数用）
' = 引数    sTargetStr      String  [in]  検索対象文字列
' = 引数    sSearchPattern  String  [in]  検索パターン
' = 引数    oMatchResult    Object  [out] 検索結果
' = 戻値                    Boolean       ヒット有無
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function ExecRegExp( _
    ByVal sTargetStr As String, _
    ByVal sSearchPattern As String, _
    ByRef oMatchResult As Object, _
    Optional ByVal bIgnoreCase As Boolean = True, _
    Optional ByVal bGlobal As Boolean = True _
) As Boolean
    Dim oRegExp As Object
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.IgnoreCase = bIgnoreCase
    oRegExp.Global = bGlobal
    oRegExp.Pattern = sSearchPattern
    Set oMatchResult = oRegExp.Execute(sTargetStr)
    If oMatchResult.Count = 0 Then
        ExecRegExp = False
    Else
        ExecRegExp = True
    End If
End Function
    Private Sub Test_ExecRegExp()
        Dim sTargetStr As String
        Dim oMatchResult As Object
        sTargetStr = "void TestFunc(int arg1, char arg2);"
        Debug.Print "*** test start! ***"
        Debug.Print ExecRegExp(sTargetStr, " \w+\(", oMatchResult)
        Debug.Print "*** test finished! ***"
    End Sub

' ==================================================================
' = 概要    クリップボードにテキストをコピー（Win32Apiを使用）
' = 引数    sText       String  [in]  コピー対象文字列
' = 戻値                Boolean       コピー結果
' = 覚書    Win32APIを使用する。
' =         ※ クリップボードは DataObject の PutInClipboard でも利用
' =            可能｡しかし､DataObject は参照設定が必要なうえ､特定のク
' =            リップボード形式には貼り付けされない｡（CF_UNICODETEXT
' =            のみで CF_TEXTへは貼り付けされない）
' =            上記のように DataObject を使用したくない場合に本関数
' =            を利用すること｡
' = 依存    user32/OpenClipboard()
' =         user32/EmptyClipboard()
' =         user32/CloseClipboard()
' =         user32/SetClipboardData()
' =         kernel32/GlobalAlloc()
' =         kernel32/GlobalLock()
' =         kernel32/GlobalUnlock()
' =         kernel32/lstrcpy()
' = 所属    Mng_Clipboard.bas
' ==================================================================
Public Function SetToClipboard( _
    sText As String _
) As Boolean
    '定数宣言
    Const GMEM_MOVEABLE         As Long = &H2
    Const GMEM_ZEROINIT         As Long = &H40
    Const GHND                  As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
    Const CF_TEXT               As Long = 1
    Const CF_OEMTEXT            As Long = 7
    
    Dim hGlobal As Long
    Dim lTextLen As Long
    Dim p As Long
    
    '戻り値をとりあえず、Falseに設定しておく。
    If OpenClipboard(0) <> 0 Then
        If EmptyClipboard() <> 0 Then
            lTextLen = LenB(sText) + 1 '長さの算出(本来はUnicodeから変換後の長さを使うほうがよい)
            hGlobal = GlobalAlloc(GHND, lTextLen) 'コピー先の領域確保
            p = GlobalLock(hGlobal)
            Call lstrcpy(p, sText) '文字列をコピー
            Call GlobalUnlock(hGlobal) 'クリップボードに渡すときにはUnlockしておく必要がある
            Call SetClipboardData(CF_TEXT, hGlobal) 'クリップボードへ貼り付ける
            Call CloseClipboard 'クリップボードをクローズ
            SetToClipboard = True 'コピー成功
        Else
            SetToClipboard = False
        End If
    Else
        SetToClipboard = False
    End If
End Function

' ==================================================================
' = 概要    アクティブブックの指定シートへ移動する
' = 引数    sSheetName      String  [in]  移動対象シート名
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Public Function JumpToTrgtSheet( _
    ByVal sSheetName As String _
)
    Dim shSheet As Worksheet
    For Each shSheet In ActiveWorkbook.Sheets
        If shSheet.Name = sSheetName Then
            shSheet.Activate
        End If
    Next
End Function


