Attribute VB_Name = "Macros"
Option Explicit

' my excel addin macros v2.24

' =============================================================================
' =  <<マクロ一覧>>
' =     ・共通
' =         F1ヘルプ無効化                              F1ヘルプを無効化する
' =
' =     ・マクロ設定
' =         マクロショートカットキー全て有効化          マクロショートカットキー全て有効化
' =         マクロショートカットキー全て無効化          マクロショートカットキー全て無効化
' =         アドインマクロ実行                          アドインマクロ実行
' =         モジュール一括エクスポート_アドイン         本アドイン内の全マクロ/プロシージャをエクスポートする
' =         モジュール一括エクスポート_アクティブブック アクティブブック内の全マクロ/プロシージャをエクスポートする
' =         CtrlShiftFマクロ                            ショートカットキー重複時の振り分け処理(Ctrl + Shift + F)
' =
' =     ・ブック操作
' =         別プロセスで開く                            アクティブブックを別プロセスで開く
' =         ファイルパスコピー                          アクティブブックのファイルパスをコピー
' =         ファイル名コピー                            アクティブブックのファイル名をコピー
' =
' =     ・シート操作
' =         EpTreeの関数ツリーをExcelで取り込む         EpTreeの関数ツリーをExcelで取り込む
' =         Excel方眼紙                                 Excel方眼紙
' =         選択シート切り出し                          選択シートを別ファイルに切り出す
' =         全シート名をコピー                          ブック内のシート名を全てコピーする
' =         シート表示非表示を切り替え                  シート表示/非表示を切り替える
' =         シート並べ替え作業用シートを作成            シート並べ替え作業用シート作成
' =         シート選択ウィンドウを表示                  シート選択ウィンドウを表示する
' =         シート名一括変更                            シート名を一括変更する
' =         シート追加カスタム                          シートを追加する（カスタム設定版）
' =         先頭シートへジャンプ                        アクティブブックの先頭シートへ移動する
' =         末尾シートへジャンプ                        アクティブブックの末尾シートへ移動する
' =         シート再計算時間計測                        シート毎に再計算にかかる時間を計測する
' =
' =     ・セル操作
' =         ファイルエクスポート                        選択範囲をファイルとしてエクスポートする。
' =         DOSコマンドを一括実行                       選択範囲内のDOSコマンドをまとめて実行する。
' =         DOSコマンドを各々実行                       選択範囲内のDOSコマンドをそれぞれ実行する。
' =         DOSコマンドを一括実行_管理者権限            選択範囲内のDOSコマンドをまとめて実行する。（管理者権限）
' =         検索文字の文字色を変更                      選択範囲内の検索文字の文字色を変更する
' =         セル内の丸数字をデクリメント                ②～⑮を指定して、指定番号以降をインクリメントする
' =         セル内の丸数字をインクリメント              ①～⑭を指定して、指定番号以降をデクリメントする
' =         ツリーをグループ化                          ツリーグループ化する
' =         ハイパーリンク一括オープン                  選択した範囲のハイパーリンクを一括で開く
' =         ハイパーリンクで飛ぶ                        アクティブセルからハイパーリンク先に飛ぶ
' =         選択範囲内で中央                            選択セルに対して「選択範囲内で中央」を実行する
' =         範囲を維持したままセルコピー                選択範囲を範囲を維持したままセルコピーする。(ダブルクオーテーションを除く)
' =         一行にまとめてセルコピー                    選択範囲を一行にまとめてセルコピーする。
' =         ●設定変更●一行にまとめてセルコピー        一行にまとめてセルコピーにて使用する「先頭文字,区切り文字,末尾文字」を変更する
' =         クリップボード値貼り付け                    クリップボードから値貼り付けする
' =         フォント色をトグル                          フォント色を「設定色」⇔「自動」でトグルする
' =         ●設定変更●フォント色をトグルの色選択      「フォント色をトグル」の設定色をカラーパレットから取得して変更する
' =         ●設定変更●フォント色をトグルの色スポイト  「フォント色をトグル」の設定色をアクティブセルから取得して変更する
' =         背景色をトグル                              背景色を「設定色」⇔「背景色なし」でトグルする
' =         ●設定変更●背景色をトグルの色選択          「背景色をトグル」の設定色をカラーパレットから取得して変更する
' =         ●設定変更●背景色をトグルの色スポイト      「背景色をトグル」の設定色をアクティブセルから取得して変更する
' =         オートフィル実行                            オートフィルを実行する
' =         画面を上に移動                              画面を上に移動(スクロールロック動作)
' =         画面を下に移動                              画面を下に移動(スクロールロック動作)
' =         画面を左に移動                              画面を左に移動(スクロールロック動作)
' =         画面を右に移動                              画面を右に移動(スクロールロック動作)
' =         インデントを上げる                          インデントを上げる
' =         インデントを下げる                          インデントを下げる
' =         アクティブセルコメントのみ表示              他セルコメントを“非表示”にしてアクティブセルコメントを“表示”にする
' =         アクティブセルコメントのみ表示して下移動    下移動後、他セルコメントを“非表示”にしてアクティブセルコメントを“表示”にする
' =         アクティブセルコメントのみ表示して上移動    上移動後、他セルコメントを“非表示”にしてアクティブセルコメントを“表示”にする
' =         アクティブセルコメントのみ表示して右移動    右移動後、他セルコメントを“非表示”にしてアクティブセルコメントを“表示”にする
' =         アクティブセルコメントのみ表示して左移動    左移動後、他セルコメントを“非表示”にしてアクティブセルコメントを“表示”にする
' =         ●設定変更●アクティブセルコメントのみ表示  アクティブセルコメント設定を切り替える
' =         Excel数式整形化実施                         Excel数式整形化実施
' =         Excel数式整形化解除                         Excel数式整形化解除
' =         セルコメントの書式設定を一括変更            セルコメントの書式設定を一括変更
' =         Diff色付け                                  選択範囲のDiff形式のフォント色に変更する。(旧:赤、新:緑)
' =         選択範囲アドレス結合文字列コピー_XXX        選択範囲のセルアドレスを結合して文字列コピー
' =
' =     ・オブジェクト操作
' =         最前面へ移動                                最前面へ移動する
' =         最背面へ移動                                最背面へ移動する
' =         オブジェクトサイズ変更プロパティ一括変更    現在シートのを全オブジェクトを
' =                                                     「セルに合わせて移動とサイズ変更をする」に変更
' =============================================================================

'******************************************************************************
'* 事前処理
'******************************************************************************
'Win32API宣言
'▽▽▽Macro.bas/範囲を維持したままセルコピー()▽▽▽
'▽▽▽Macro.bas/一行にまとめてセルコピー()▽▽▽
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'△△△Macro.bas/一行にまとめてセルコピー()△△△
'△△△Macro.bas/範囲を維持したままセルコピー()△△△

'▽▽▽Mng_Clipboard.bas/SetToClipboard()▽▽▽
'▽▽▽Mng_Clipboard.bas/GetFromClipboard()▽▽▽
#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
#Else
Private Declare Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#End If
Private Const GHND = &H42
Private Const CF_TEXT = &H1
Private Const CF_LINK = &HBF00
Private Const CF_BITMAP = 2
Private Const CF_METAFILE = 3
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const MAXSIZE = 4096
'△△△Mng_Clipboard.bas/GetFromClipboard()△△△
'△△△Mng_Clipboard.bas/SetToClipboard()△△△

'▽▽▽Macro.bas/ShowColorPalette()▽▽▽
Private Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'△△△Macro.bas/ShowColorPalette()△△△

'▽▽▽Macro.bas/ReadSettingFile()/WriteSettingFile()▽▽▽
Const sDELIMITER_INIT As String = vbTab
'△△△Macro.bas/ReadSettingFile()/WriteSettingFile()△△△

'▽▽▽Mng_SysCmd.bas/ExecDosCmdRunas()▽▽▽
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long
'△△△Mng_SysCmd.bas/ExecDosCmdRunas()△△△

'******************************************************************************
'* 設定値
'******************************************************************************
'▼▼▼ 設定(初期値) ▼▼▼
'=== 背景色をトグル()/フォント色をトグル() ===
    '[色名参考] https://excel-toshokan.com/vba-color-list/
    Const lCLRTGLBG_CLR_RGB As Long = vbYellow
    Const lCLRTGLFONT_CLR_RGB As Long = vbRed
'=== アクティブセルコメント設定() ===
    Const bCMNT_VSBL_ENB As Boolean = False
'=== Excel方眼紙() ===
    Const sEXCELGRID_FONT_NAME As String = "ＭＳ ゴシック"
    Const lEXCELGRID_FONT_SIZE As Long = 9
    Const lEXCELGRID_CLM_WIDTH As Long = 3 '3文字分
'=== 検索文字の文字色を変更() ===
    Const sWORDCOLOR_SRCH_WORD As String = ""
    Const lWORDCOLOR_CLR_RGB As Long = vbRed
'=== ファイルエクスポート() ===
    Const sFILEEXPORT_OUT_FILE_NAME As String = "MyExcelAddinFileExport.csv"
    Const bFILEEXPORT_IGNORE_INVISIBLE_CELL As Boolean = True
    Const sFILEEXPORT_CHAR_SET As String = "Shift_JIS" '(UTF-8|UTF-16|Shift_JIS|EUC-JP|ISO-2022-JP|...)
    Const lFILEEXPORT_LINE_SEPARATER As Long = 10 '13:CR 10:LF -1:CRLF
'=== DOSコマンドを一括実行() ===
    Const sCMDEXEBAT_REDIRECT_FILE_NAME As String = "MyExcelAddinCmdExeBat.log"
    Const sCMDEXEBAT_BAT_FILE_NAME As String = "MyExcelAddinCmdExeBat.bat"
    Const bCMDEXEBAT_IGNORE_INVISIBLE_CELL As Boolean = True
'=== DOSコマンドを一括実行_管理者権限() ===
    Const sCMDEXEBATRUNAS_REDIRECT_FILE_NAME As String = "MyExcelAddinCmdExeBatRunas.log"
    Const bCMDEXEBATRUNAS_IGNORE_INVISIBLE_CELL As Boolean = True
'=== DOSコマンドを各々実行() ===
    Const sCMDEXEUNI_REDIRECT_FILE_NAME As String = "MyExcelAddinCmdExeUni.log"
    Const bCMDEXEUNI_IGNORE_INVISIBLE_CELL As Boolean = True
'=== EpTreeの関数ツリーをExcelで取り込む() ===
    Const sEPTREE_OUT_SHEET_NAME As String = "CallTree"
    Const lEPTREE_MAX_FUNC_LEVEL_INI As Long = 10
    Const lEPTREE_CLM_WIDTH As Long = 2
    Const sEPTREE_OUT_LOG_PATH As String = "c:\"
    Const sEPTREE_DEV_ROOT_DIR_PATH As String = "c:\"
    Const lEPTREE_DEV_ROOT_DIR_LEVEL As Long = 0
'=== 範囲を維持したままセルコピー() ===
    Const bCELLCOPYRNG_IGNORE_INVISIBLE_CELL As Boolean = True
    Const sCELLCOPYRNG_DELIMITER As String = vbTab
'=== 一行にまとめてセルコピー() ===
    Const bCELLCOPYLINE_IGNORE_INVISIBLE_CELL As Boolean = True
    Const bCELLCOPYLINE_IGNORE_BLANK_CELL As Boolean = True
    Const sCELLCOPYLINE_PREFFIX As String = "("
    Const sCELLCOPYLINE_DELIMITER As String = "|"
    Const sCELLCOPYLINE_SUFFIX As String = ")"
'=== シート選択ウィンドウを表示() ===
    Const bSHTSELWIN_MSGBOX_SHOW As Boolean = False
'=== 選択範囲アドレス結合文字列コピー_xxx() ===
    Const sCELLADRJOIN_DELIMITER As String = ""
    Const bCELLADRJOIN_FORMAT_R1C1 As Boolean = False
'▲▲▲ 設定 ▲▲▲

' ==================================================================
' = 概要    ショートカットキーの有効/無効を切り替える
' = 引数    bActivateShortcutKeys   Boolean     [in]    有効化/無効化
' = 覚書    なし
' = 依存    SettingFile.cls
' = 所属    Macros.bas
' ==================================================================
Private Sub SwitchMacroShortcutKeysActivation( _
    ByVal bActivateShortcutKeys As Boolean _
)
    Dim dMacroShortcutKeys As Object
    Set dMacroShortcutKeys = CreateObject("Scripting.Dictionary")
    
    '*** アドイン設定読み出し ***
    Dim bCmntVsblEnb As Boolean
    bCmntVsblEnb = ReadSettingFile("bCMNT_VSBL_ENB", bCMNT_VSBL_ENB)
    
    '*** ショートカットキー設定更新 ***
    ' <<ショートカットキー追加方法>>
    '   dMacroShortcutKeysに対してキー<マクロ名>、値<ショートカットキー>を追加する。
    '   第一引数にはショートカットキー、第二引数にマクロ名を指定する。
    '   ショートカットキーは Ctrl や Shift などと組み合わせて指定できる。
    '     Ctrl：^、Shift：+、Alt：%
    '   詳細は以下 URL 参照。
    '     https://msdn.microsoft.com/ja-jp/library/office/ff197461.aspx
    '▼▼▼ 設定 ▼▼▼
    '共通
'   dMacroShortcutKeys.Add "", "F1ヘルプ無効化"
    
    'マクロ設定
'   dMacroShortcutKeys.Add "", "マクロショートカットキー全て有効化"
'   dMacroShortcutKeys.Add "", "マクロショートカットキー全て無効化"
    dMacroShortcutKeys.Add "+%{F8}", "アドインマクロ実行"
'   dMacroShortcutKeys.Add "", "モジュール一括エクスポート_アドイン"
'   dMacroShortcutKeys.Add "", "モジュール一括エクスポート_アクティブブック"
    dMacroShortcutKeys.Add "^+f", "CtrlShiftFマクロ"
    
    'ブック操作
'   dMacroShortcutKeys.Add "", "別プロセスで開く"
    dMacroShortcutKeys.Add "^%p", "ファイルパスコピー"
    dMacroShortcutKeys.Add "^%n", "ファイル名コピー"
    
    'シート操作
'   dMacroShortcutKeys.Add "", "EpTreeの関数ツリーをExcelで取り込む"
    dMacroShortcutKeys.Add "^%h", "Excel方眼紙"
'   dMacroShortcutKeys.Add "", "選択シート切り出し"
    dMacroShortcutKeys.Add "^%c", "全シート名をコピー"
'   dMacroShortcutKeys.Add "", "シート表示非表示を切り替え"
'   dMacroShortcutKeys.Add "", "シート並べ替え作業用シートを作成"
    dMacroShortcutKeys.Add "^%{PGUP}", "シート選択ウィンドウを表示"
    dMacroShortcutKeys.Add "^%{PGDN}", "シート選択ウィンドウを表示"
'   dMacroShortcutKeys.Add "", "シート名一括変更"
    dMacroShortcutKeys.Add "+{F11}", "シート追加カスタム"
    dMacroShortcutKeys.Add "^%{HOME}", "先頭シートへジャンプ"
    dMacroShortcutKeys.Add "^%{END}", "末尾シートへジャンプ"
'   dMacroShortcutKeys.Add "", "シート再計算時間計測"
    
    'セル操作
'   dMacroShortcutKeys.Add "", "ファイルエクスポート"
'   dMacroShortcutKeys.Add "", "DOSコマンドを一括実行"
'   dMacroShortcutKeys.Add "", "DOSコマンドを各々実行"
'   dMacroShortcutKeys.Add "", "DOSコマンドを一括実行_管理者権限"
'   dMacroShortcutKeys.Add "^+f", "検索文字の文字色を変更" '「CtrlShiftFマクロ」にて実行
'   dMacroShortcutKeys.Add "", "セル内の丸数字をデクリメント"
'   dMacroShortcutKeys.Add "", "セル内の丸数字をインクリメント"
'   dMacroShortcutKeys.Add "", "ツリーをグループ化"
'   dMacroShortcutKeys.Add "", "選択範囲のセルアドレスを結合して文字列コピー"
'   dMacroShortcutKeys.Add "", "ハイパーリンク一括オープン"
    dMacroShortcutKeys.Add "^+j", "ハイパーリンクで飛ぶ"
'   dMacroShortcutKeys.Add "", "選択範囲内で中央"
    dMacroShortcutKeys.Add "^+c", "範囲を維持したままセルコピー"
    dMacroShortcutKeys.Add "^+d", "一行にまとめてセルコピー"
    dMacroShortcutKeys.Add "^%d", "●設定変更●一行にまとめてセルコピー"
'   dMacroShortcutKeys.Add "^+v", "クリップボード値貼り付け" 'マクロ使用時はアンドゥできないため、極力使用しない
    dMacroShortcutKeys.Add "^2", "背景色をトグル"
    dMacroShortcutKeys.Add "^%2", "●設定変更●背景色をトグルの色選択"
    dMacroShortcutKeys.Add "+%2", "●設定変更●背景色をトグルの色スポイト"
    dMacroShortcutKeys.Add "^3", "フォント色をトグル"
    dMacroShortcutKeys.Add "^%3", "●設定変更●フォント色をトグルの色選択"
    dMacroShortcutKeys.Add "+%3", "●設定変更●フォント色をトグルの色スポイト"
'   dMacroShortcutKeys.Add "^%{DOWN}", "'オートフィル実行(""Down"")'"
'   dMacroShortcutKeys.Add "^%{UP}", "'オートフィル実行(""Up"")'"
    dMacroShortcutKeys.Add "^%{UP}", "画面を上に移動"
    dMacroShortcutKeys.Add "^%{DOWN}", "画面を下に移動"
    dMacroShortcutKeys.Add "^%{LEFT}", "画面を左に移動"
    dMacroShortcutKeys.Add "^%{RIGHT}", "画面を右に移動"
    dMacroShortcutKeys.Add "^+>", "インデントを上げる"
    dMacroShortcutKeys.Add "^+<", "インデントを下げる"
    If bCmntVsblEnb = True Then
        dMacroShortcutKeys.Add "{DOWN}", "アクティブセルコメントのみ表示して下移動"
        dMacroShortcutKeys.Add "{UP}", "アクティブセルコメントのみ表示して上移動"
        dMacroShortcutKeys.Add "{RIGHT}", "アクティブセルコメントのみ表示して右移動"
        dMacroShortcutKeys.Add "{LEFT}", "アクティブセルコメントのみ表示して左移動"
    Else
        dMacroShortcutKeys.Add "{DOWN}", ""
        dMacroShortcutKeys.Add "{UP}", ""
        dMacroShortcutKeys.Add "{RIGHT}", ""
        dMacroShortcutKeys.Add "{LEFT}", ""
    End If
    dMacroShortcutKeys.Add "^+{F11}", "●設定変更●アクティブセルコメントのみ表示"
    dMacroShortcutKeys.Add "^+i", "Excel数式整形化実施"
    dMacroShortcutKeys.Add "^%i", "Excel数式整形化解除"
'   dMacroShortcutKeys.Add "", "セルコメントの書式設定を一括変更"
    dMacroShortcutKeys.Add "+%d", "Diff色付け"
    
    'オブジェクト操作
'   dMacroShortcutKeys.Add "^+f", "最前面へ移動" '「CtrlShiftFマクロ」にて実行
    dMacroShortcutKeys.Add "^+b", "最背面へ移動"
    '▲▲▲ 設定 ▲▲▲
    
    '*** ショートカットキー設定反映 ***
    Dim vShortcutKey As Variant
    Dim sMacroName As String
    If bActivateShortcutKeys = True Then
        For Each vShortcutKey In dMacroShortcutKeys
            sMacroName = dMacroShortcutKeys.Item(vShortcutKey)
            If sMacroName = "" Then
                Application.OnKey CStr(vShortcutKey)              'ショートカットキークリア
            Else
                Application.OnKey CStr(vShortcutKey), sMacroName  'ショートカットキー設定
            End If
        Next
    Else
        For Each vShortcutKey In dMacroShortcutKeys
            Application.OnKey CStr(vShortcutKey)                  'ショートカットキークリア
        Next
    End If
End Sub

' *****************************************************************************
' * 外部公開用マクロ
' *****************************************************************************
Private Sub ▼▼▼▼▼外部公開用マクロ▼▼▼▼▼()
    'プロシージャリスト表示用のダミープロシージャ
End Sub

' ▽▽▽ 共通 ▽▽▽
' =============================================================================
' = 概要    F1ヘルプを無効化する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub F1ヘルプ無効化()
    Application.OnKey "{F1}", ""
End Sub

' ▽▽▽ マクロ設定 ▽▽▽
' =============================================================================
' = 概要    マクロショートカットキー全て有効化
' = 覚書    なし
' = 依存    Macros.bas/SwitchMacroShortcutKeysActivation()
' = 所属    Macros.bas
' =============================================================================
Public Sub マクロショートカットキー全て有効化()
    Call SwitchMacroShortcutKeysActivation(True)
    
    Application.StatusBar = "■■■マクロショートカットキーを有効化しました■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' =============================================================================
' = 概要    マクロショートカットキー全て無効化
' = 覚書    なし
' = 依存    Macros.bas/SwitchMacroShortcutKeysActivation()
' = 所属    Macros.bas
' =============================================================================
Public Sub マクロショートカットキー全て無効化()
    Call SwitchMacroShortcutKeysActivation(False)
    
    Application.StatusBar = "■■■マクロショートカットキーを無効化しました■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' =============================================================================
' = 概要    アドインマクロ実行
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub アドインマクロ実行()
    ExecAddInMacro.Show
End Sub

' =============================================================================
' = 概要    本アドイン内の全マクロ/プロシージャをエクスポートする
' = 覚書    ・以下の参照設定を追加する必要あり。
' =           - [ツール] -> [参照設定] ->「Microsoft Visual Basic for Applications Extensibility」
' = 依存    Macros.bas/ExportAllModules()
' = 所属    Macros.bas
' =============================================================================
Public Sub モジュール一括エクスポート_アドイン()
    Const sMACRO_NAME As String = "モジュール一括エクスポート_アドイン"
    Call ExportAllModules(ThisWorkbook)
    MsgBox "アドイン内の全モジュールをエクスポートしました！", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = 概要    アクティブブック内の全マクロ/プロシージャをエクスポートする
' = 覚書    ・以下の参照設定を追加する必要あり。
' =           - [ツール] -> [参照設定] ->「Microsoft Visual Basic for Applications Extensibility」
' = 依存    Macros.bas/ExportAllModules()
' = 所属    Macros.bas
' =============================================================================
Public Sub モジュール一括エクスポート_アクティブブック()
    Const sMACRO_NAME As String = "モジュール一括エクスポート_アクティブブック"
    Call ExportAllModules(ActiveWorkbook)
    MsgBox "アクティブブック内の全モジュールをエクスポートしました！", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = 概要    ショートカットキー重複時の振り分け処理（Ctrl+Shift+F）
' = 覚書    なし
' = 依存    Macros.bas/最前面へ移動()
' =         Macros.bas/検索文字の文字色を変更()
' = 所属    Macros.bas
' =============================================================================
Public Sub CtrlShiftFマクロ()
    On Error Resume Next
    Call 最前面へ移動
    If Err.Number <> 0 Then
        Call 検索文字の文字色を変更
    End If
    On Error GoTo 0
End Sub

' ▽▽▽ ブック操作 ▽▽▽
' =============================================================================
' = 概要    アクティブブックを別プロセスで開く
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 別プロセスで開く()
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sActiveBookPath
    sActiveBookPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    objWshShell.Run "cmd /c excel /x /r """ & sActiveBookPath & """", 0, False
End Sub

' =============================================================================
' = 概要    アクティブブックのファイルパスをコピー
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub ファイルパスコピー()
    Const sMACRO_NAME As String = "ファイルパスコピー"
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
        .PutInClipboard
    End With
    
    '*** フィードバック ***
    Application.StatusBar = "■■■■■■■■ " & sMACRO_NAME & "完了！ ■■■■■■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' =============================================================================
' = 概要    アクティブブックのファイル名をコピー
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub ファイル名コピー()
    Const sMACRO_NAME As String = "ファイル名コピー"
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText ActiveWorkbook.Name
        .PutInClipboard
    End With
    
    '*** フィードバック ***
    Application.StatusBar = "■■■■■■■■ " & sMACRO_NAME & "完了！ ■■■■■■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' ▽▽▽ シート操作 ▽▽▽
' =============================================================================
' = 概要    EpTreeの関数ツリーをExcelで取り込む
' = 覚書    なし
' = 依存    Mng_FileSys.bas/ShowFileSelectDialog()
' =         Mng_FileSys.bas/ShowFolderSelectDialog()
' =         Mng_Collection.bas/ReadTxtFileToCollection()
' =         Mng_String.bas/ExecRegExp()
' =         Mng_String.bas/ExtractTailWord()
' =         Mng_String.bas/ExtractRelativePath()
' =         Mng_ExcelOpe.bas/CreateNewWorksheet()
' =         SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub EpTreeの関数ツリーをExcelで取り込む()
    Const sMACRO_NAME As String = "EpTreeの関数ツリーをExcelで取り込む"
    Const lSTRT_ROW As Long = 1
    Const lSTRT_CLM As Long = 1
    
    Dim lRowIdx As Long
    Dim lStrtRow As Long
    Dim lLastRow As Long
    Dim lStrtClm As Long
    Dim lLastClm As Long
    
    '=============================================
    '= 事前処理
    '=============================================
    Dim sOutSheetName As String
    Dim lMaxFuncLevelIni As Long
    Dim lClmWidth As Long
    Dim sEptreeLogPath As String
    Dim sDevRootDirPath As String
    Dim sDevRootDirName As String
    Dim lDevRootLevel As Long
    
    '*** アドイン設定ファイルから設定読み出し ***
    sOutSheetName = ReadSettingFile("sEPTREE_OUT_SHEET_NAME", sEPTREE_OUT_SHEET_NAME)
    lMaxFuncLevelIni = ReadSettingFile("lEPTREE_MAX_FUNC_LEVEL_INI", lEPTREE_MAX_FUNC_LEVEL_INI)
    lClmWidth = ReadSettingFile("lEPTREE_CLM_WIDTH", lEPTREE_CLM_WIDTH)
    
    'Eptreeログファイルパス取得
    sEptreeLogPath = ReadSettingFile("sEPTREE_OUT_LOG_PATH", sEPTREE_OUT_LOG_PATH)
    sEptreeLogPath = ShowFileSelectDialog(sEptreeLogPath, "EpTreeLog.txtのファイルパスを選択してください")
    If sEptreeLogPath = "" Then
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    Call WriteSettingFile("sEPTREE_OUT_LOG_PATH", sEptreeLogPath)
    
    '開発用ルートフォルダ取得
    sDevRootDirPath = ReadSettingFile("sEPTREE_DEV_ROOT_DIR_PATH", sEPTREE_DEV_ROOT_DIR_PATH)
    sDevRootDirPath = ShowFolderSelectDialog(sDevRootDirPath, "開発用ルートフォルダパスを選択してください（空欄の場合は親フォルダが選択されます）")
    If sDevRootDirPath = "" Then
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    sDevRootDirName = ExtractTailWord(sDevRootDirPath, "\")
    Call WriteSettingFile("sEPTREE_DEV_ROOT_DIR_PATH", sDevRootDirPath)
    
    'ルートフォルダレベル取得
    lDevRootLevel = ReadSettingFile("lEPTREE_DEV_ROOT_DIR_LEVEL", lEPTREE_DEV_ROOT_DIR_LEVEL)
    Dim sDevRootLevel As String
    sDevRootLevel = InputBox("ルートフォルダレベルを入力してください", sMACRO_NAME, CStr(lDevRootLevel))
    If sDevRootLevel = "" Then
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    Call WriteSettingFile("lEPTREE_DEV_ROOT_DIR_LEVEL", sDevRootLevel)
    
    '=============================================
    '= 本処理
    '=============================================
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'シート追加
    Dim sSheetName As String
    Dim shTrgtSht As Worksheet
    sSheetName = CreateNewWorksheet(sOutSheetName)
    Set shTrgtSht = ActiveWorkbook.Sheets(sSheetName)
    
    'テキストファイル読み出し
    Dim cFileContents As Collection
    Set cFileContents = New Collection
    Call ReadTxtFileToCollection(sEptreeLogPath, cFileContents)
    
    'ファイルツリー出力
    lStrtRow = lSTRT_ROW
    lStrtClm = lSTRT_CLM
    lRowIdx = lStrtRow
    
    With shTrgtSht
        .Cells(lRowIdx, lStrtClm + 0).Value = "ファイルパス"
        .Cells(lRowIdx, lStrtClm + 1).Value = "行数"
        .Cells(lRowIdx, lStrtClm + 2).Value = "関数名"
        .Cells(lRowIdx, lStrtClm + 3).Value = "コールツリー"
    End With
    lRowIdx = lRowIdx + 1
    
    Dim lMaxFuncLevel As Long
    lMaxFuncLevel = lMaxFuncLevelIni
    Dim vFileLine As Variant
    For Each vFileLine In cFileContents
        Dim oMatchResult As Object
        Call ExecRegExp( _
            vFileLine, _
            "^([^ ]+)? +(\d+): (  )?([│|└|├|  ]*)(\w+)(↑)?", _
            oMatchResult _
        )
        
        Dim sFilePath As String
        Dim sLineNo As String
        Dim lFuncLevel As Long
        Dim sFuncName As String
        Dim sOmission As String
        sFilePath = oMatchResult(0).SubMatches(0)
        Call ExtractRelativePath(sFilePath, sDevRootDirName, Int(sDevRootLevel), sFilePath)
        sLineNo = oMatchResult(0).SubMatches(1)
        If sLineNo = 0 Then
            sLineNo = ""
        End If
        lFuncLevel = LenB(StrConv(oMatchResult(0).SubMatches(3), vbFromUnicode)) / 2
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
        lLastClm = lSTRT_CLM + 3 + lMaxFuncLevel
        lLastRow = lRowIdx
        
        'タイトル行 中央揃え
        .Range(.Cells(lStrtRow, lStrtClm + 0), .Cells(lStrtRow, lStrtClm + 2)).HorizontalAlignment = xlCenter
        .Range(.Cells(lStrtRow, lStrtClm + 3), .Cells(lStrtRow, lLastClm)).HorizontalAlignment = xlCenterAcrossSelection
        
        '列幅調整
        .Range(.Cells(lStrtRow, lStrtClm + 0), .Cells(lLastRow, lStrtClm + 0)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 1), .Cells(lLastRow, lStrtClm + 1)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 2), .Cells(lLastRow, lStrtClm + 2)).Columns.AutoFit
        .Range(.Cells(lStrtRow, lStrtClm + 3), .Cells(lLastRow, lLastClm)).ColumnWidth = lClmWidth
        
        'オートフィルタ
        .Range(.Cells(lStrtRow, lStrtClm), .Cells(lLastRow, lLastClm)).AutoFilter
        
        '行高さ
        .Rows(lStrtRow).RowHeight = .Rows(lStrtRow).RowHeight * 3
        
        'タイトル列固定
        ActiveWindow.FreezePanes = False
        .Rows(lStrtRow + 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1).Select
        
        'シート見出し色
        .Tab.Color = RGB(242, 220, 219)
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "関数コールツリー作成完了！", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = 概要    Excel方眼紙
' = 覚書    なし
' = 依存    SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub Excel方眼紙()
    'アドイン設定読み出し
    Dim sFontName As String
    Dim lFontSize As Long
    Dim lClmWidth As Long
    sFontName = ReadSettingFile("sEXCELGRID_FONT_NAME", sEXCELGRID_FONT_NAME)
    lFontSize = ReadSettingFile("lEXCELGRID_FONT_SIZE", lEXCELGRID_FONT_SIZE)
    lClmWidth = ReadSettingFile("lEXCELGRID_CLM_WIDTH", lEXCELGRID_CLM_WIDTH)
    
    'Excel方眼紙設定
    ActiveSheet.Cells.Select
    With Selection
        .Font.Name = sFontName
        .Font.Size = lFontSize
        .ColumnWidth = lClmWidth
        .Rows.AutoFit
    End With
    ActiveSheet.Cells(1, 1).Select
End Sub

' ==================================================================
' = 概要    選択シートを別ファイルに切り出す。
' =         コピー元ブックと同フォルダに出力する。
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Public Sub 選択シート切り出し()
    Const sMACRO_NAME As String = "選択シート切り出し"
    
    Dim shSht As Worksheet
    Dim wSrcWindow As Window
    Dim bkSrcBook As Workbook
    Dim bkTrgtBook As Workbook
    Dim sTrgtBookName As String
    
    Set bkSrcBook = ActiveWorkbook
    Set wSrcWindow = ActiveWindow
    Set bkTrgtBook = Workbooks.Add
    
    wSrcWindow.SelectedSheets.Copy _
        After:=bkTrgtBook.Sheets(bkTrgtBook.Sheets.Count)
    Application.DisplayAlerts = False
    bkTrgtBook.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    bkTrgtBook.SaveAs bkSrcBook.Path & "\" & wSrcWindow.SelectedSheets(1).Name & ".xlsx"
    bkTrgtBook.Close
    
    MsgBox "選択シート切り出し完了！", vbOKOnly, sMACRO_NAME
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
    Const sMACRO_NAME As String = "全シート名をコピー"
    
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
    
    MsgBox "ブック内のシート名を全てコピーしました", vbOKOnly, sMACRO_NAME
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
' = 概要    シートを並び替える。
' =         本処理を実行すると、シート並べ替え作業用シートを作成する。
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub シート並べ替え作業用シートを作成()
    Const sMACRO_NAME As String = "シート並べ替え作業用シートを作成"
    Const WORK_SHEET_NAME As String = "シート並べ替え作業用"
    Const ROW_BTN = 2
    Const ROW_TEXT_1 = 4
    Const ROW_TEXT_2 = 5
    Const ROW_SHT_NAME_TITLE = 7
    Const ROW_SHT_NAME_STRT = 8
    Const CLM_BTN = 2
    Const CLM_SHT_NAME = 2
    
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
            MsgBox "既に「" & WORK_SHEET_NAME & "」シートが作成されています。", vbCritical, sMACRO_NAME
            MsgBox "処理を続けたい場合は、シートを削除してください。", vbCritical, sMACRO_NAME
            MsgBox "処理を中断します。", vbCritical, sMACRO_NAME
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
' = 概要    シート選択ウィンドウを表示する
' = 覚書    ・環境によってはシートがアクティブ化されないことがあるが、
' =           なぜか事前にMsgBoxすれば対処できる。
' =         ・bUseMyUserForm = Trueにより、自作ユーザフォームのシート
' =           選択ウィンドウを表示できる｡
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub シート選択ウィンドウを表示()
    Const bUseMyUserForm As Boolean = True
    If bUseMyUserForm = True Then
        SelectActivationSheet.Show
    Else
        Dim bMsgBoxShow As Boolean
        bMsgBoxShow = ReadSettingFile("bSHTSELWIN_MSGBOX_SHOW", bSHTSELWIN_MSGBOX_SHOW)
        If bMsgBoxShow = True Then
            MsgBox "シート選択ウィンドウを表示します", vbOKOnly, "シート選択ウィンドウ表示"
        Else
            'Do Nothing
        End If

        Application.ScreenUpdating = False
        With CommandBars.Add(Temporary:=True)
            .Controls.Add(ID:=957).Execute
            .Delete
        End With
        Application.ScreenUpdating = True
    End If
End Sub

' =============================================================================
' = 概要    シート名を一括変更する
' = 覚書    ・★要改造：2行目以降、2列目に旧シート名、3列目に新シート名を指定する。
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub シート名一括変更()
    Const lOLD_SHTNAME_CLM As Long = 2
    Const lNEW_SHTNAME_CLM As Long = 3
    Const lSTART_ROW As Long = 2
    Application.ScreenUpdating = False
    With ActiveSheet
        Dim lStrtRow As Long
        Dim lLastRow As Long
        lStrtRow = lSTART_ROW
        lLastRow = .Cells(.Rows.Count, lOLD_SHTNAME_CLM).End(xlUp).Row
        Dim lRowIdx As Long
        For lRowIdx = lStrtRow To lLastRow
            Dim sShtNameOld As String
            Dim sShtNameNew As String
            sShtNameOld = .Cells(lRowIdx, lOLD_SHTNAME_CLM).Value
            sShtNameNew = .Cells(lRowIdx, lNEW_SHTNAME_CLM).Value
            If sShtNameOld <> sShtNameNew Then
                ActiveWorkbook.Sheets(sShtNameOld).Name = sShtNameNew
            End If
        Next
    End With
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = 概要    シートを追加する（カスタム設定版）
' = 覚書    ・シート追加時、以下を実施する
' =           - アウトライン時に集計行を上、集計列を左に設定する
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub シート追加カスタム()
    'MsgBox "カスタム設定版シート追加"
    Application.ScreenUpdating = False
    Dim shAddSht As Worksheet
    Set shAddSht = ActiveWorkbook.Sheets.Add()
    shAddSht.Outline.SummaryRow = xlAbove
    shAddSht.Outline.SummaryColumn = xlLeft
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = 概要    アクティブブックの先頭シートへ移動する
' = 覚書    なし
' = 依存　　なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 先頭シートへジャンプ()
    Application.ScreenUpdating = False
    Dim shSheet As Worksheet
    For Each shSheet In ActiveWorkbook.Sheets
        If shSheet.Visible = True Then
            shSheet.Activate
            Exit For
        End If
    Next
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = 概要    アクティブブックの末尾シートへ移動する
' = 覚書    なし
' = 依存　　なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 末尾シートへジャンプ()
    Application.ScreenUpdating = False
    With ActiveWorkbook
        Dim lShtCnt As Long
        For lShtCnt = .Sheets.Count To 1 Step -1
            If .Sheets(lShtCnt).Visible = True Then
                .Sheets(lShtCnt).Activate
                Exit For
            End If
        Next
    End With
    Application.ScreenUpdating = True
End Sub

' ==================================================================
' = 概要    シート毎に再計算にかかる時間を計測する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Public Sub シート再計算時間計測()
    Const lLOOP_NUM As Long = 10
    
    Application.ScreenUpdating = False
    
    Dim previonsCalculationcMode As Variant
    previonsCalculationcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    ' 入力値が1未満の場合は終了する
    If Not (lLOOP_NUM > 0) Then
        Exit Sub
    End If
    
    ' ### ベンチマーク開始 ###
    ' ヘッダー(シート名の一覧)を出力
    Dim shTrgtSheet As Worksheet
    For Each shTrgtSheet In ActiveWorkbook.Worksheets
        Debug.Print shTrgtSheet.Name & vbTab;
    Next
    Debug.Print
    
    Dim lLoopIdx As Long
    For lLoopIdx = 1 To lLOOP_NUM
        ' シートごとに再計算&処理時間を出力
        Dim vStartTime As Variant
        Dim vFinishTime As Variant
        For Each shTrgtSheet In Worksheets
            shTrgtSheet.Cells.Dirty
            vStartTime = Timer
            shTrgtSheet.Calculate
            vFinishTime = Timer
            Debug.Print Format(vFinishTime - vStartTime, "0.0000") & vbTab;
        Next
        Debug.Print
    Next
    
    Application.Calculation = previonsCalculationcMode
    Application.ScreenUpdating = True
End Sub

' ▽▽▽ セル操作 ▽▽▽
' =============================================================================
' = 概要    選択範囲をファイルとしてエクスポートする。
' =         隣り合った列のセルにはタブ文字を挿入して出力する。
' = 覚書    なし
' = 依存    Mng_FileSys.bas/ShowFolderSelectDialog()
' =         Mng_FileSys.bas/OutputTxtFile()
' =         Mng_Array.bas/ConvRange2Array()
' =         SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub ファイルエクスポート()
    Const sMACRO_NAME As String = "ファイルエクスポート"
    
    Dim dicDelimiter As Object
    Set dicDelimiter = CreateObject("Scripting.Dictionary")
    
    '▼▼▼設定▼▼▼
    dicDelimiter.Add "csv", ","
    dicDelimiter.Add "tsv", vbTab
    '▲▲▲設定▲▲▲
    
    '*** セル選択判定 ***
    If Selection.Count = 0 Then
        MsgBox "セルが選択されていません", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        End
    End If
    
    '*** アドイン設定ファイルパス取得 ***
    '*** アドイン設定読み出し ***
    Dim bIgnoreInvisibleCell As Boolean
    bIgnoreInvisibleCell = ReadSettingFile("bFILEEXPORT_IGNORE_INVISIBLE_CELL", bFILEEXPORT_IGNORE_INVISIBLE_CELL)
    
    '*** 出力先入力 ***
    'フォルダパス
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputDirPathInit As String
    Dim sOutputDirPath As String
    sOutputDirPathInit = objWshShell.SpecialFolders("Desktop")
    sOutputDirPath = ReadSettingFile("sFILEEXPORT_OUT_DIR_PATH", sOutputDirPathInit)
    sOutputDirPath = ShowFolderSelectDialog(sOutputDirPath)
    If sOutputDirPath = "" Then
        MsgBox "無効なフォルダを指定もしくはフォルダが選択されませんでした。", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します。", vbCritical, sMACRO_NAME
        End
    Else
        'Do Nothing
    End If
    Call WriteSettingFile("sFILEEXPORT_OUT_DIR_PATH", sOutputDirPath)
    
    'ファイル名
    Dim sOutputFileName As String
    Dim sOutputFilePath As String
    Dim sFileExt As String
    Dim sDelimiter As String
    sOutputFileName = ReadSettingFile("sFILEEXPORT_OUT_FILE_NAME", sFILEEXPORT_OUT_FILE_NAME)
    sOutputFileName = InputBox("ファイル名を入力してください。(拡張子付き)", sMACRO_NAME, sOutputFileName)
    If InStr(sOutputFileName, ".") Then
        'Do Nothing
    Else
        MsgBox "ファイル名が指定されませんでした。", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します。", vbCritical, sMACRO_NAME
        End
    End If
    Call WriteSettingFile("sFILEEXPORT_OUT_FILE_NAME", sOutputFileName)
    
    'ファイルパス
    sOutputFilePath = sOutputDirPath & "\" & sOutputFileName
    
    '*** 拡張子,デリミタ取得 ***
    sFileExt = Split(sOutputFileName, ".")(UBound(Split(sOutputFileName, ".")))
    If dicDelimiter.Exists(sFileExt) Then
        sDelimiter = dicDelimiter.Item(sFileExt)
    Else
        sDelimiter = vbTab
    End If
    
    '*** ファイル上書き判定 ***
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(sOutputFilePath) Then
        Dim vAnswer As Variant
        vAnswer = MsgBox("ファイルが存在します。上書きしますか？", vbOKCancel, sMACRO_NAME)
        If vAnswer = vbOK Then
            'Do Nothing
        Else
            MsgBox "処理を中断します。", vbExclamation, sMACRO_NAME
            End
        End If
    Else
        'Do Nothing
    End If
    
    '*** Range型からString()型へ変換 ***
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIgnoreInvisibleCell, _
                sDelimiter _
            )
    
    '*** ファイル出力処理 ***
    Call OutputTxtFile(sOutputFilePath, asRange, sFILEEXPORT_CHAR_SET, lFILEEXPORT_LINE_SEPARATER)
    
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
' = 概要    選択範囲内のDOSコマンドをバッチファイルに書き出してまとめて実行する。
' =         単一列選択時のみ有効。
' = 覚書    ・大量のコマンドを実行する場合、「DOSコマンドを各々実行()」に比べて
' =           本マクロのほうが早い。
' = 依存    Mng_Array.bas/ConvRange2Array()
' =         Mng_FileSys.bas/OutputTxtFile()
' =         Mng_SysCmd.bas/ExecDosCmd()
' =         SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub DOSコマンドを一括実行()
    Const sMACRO_NAME As String = "DOSコマンドを一括実行"
    
    '*** アドイン設定読み出し ***
    Dim bIgnoreInvisibleCell As Boolean
    bIgnoreInvisibleCell = ReadSettingFile("bCMDEXEBAT_IGNORE_INVISIBLE_CELL", bCMDEXEBAT_IGNORE_INVISIBLE_CELL)
    
    '*** セル選択判定 ***
    If Selection.Count = 0 Then
        MsgBox "セルが選択されていません", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        End
    End If
    
    '*** 範囲チェック ***
    If Selection.Columns.Count = 1 Then
        'Do Nothing
    Else
        MsgBox "単一列のみ選択してください", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        End
    End If
    
    'Range型からString()型へ変換
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIgnoreInvisibleCell, _
                "" _
            )
    
    Dim sBatFileDirPath As String
    Dim sBatFilePath As String
    sBatFileDirPath = GetAddinSettingDirPath()
    sBatFilePath = sBatFileDirPath & "\" & sCMDEXEBAT_BAT_FILE_NAME
    Debug.Print sBatFilePath
    
    Call OutputTxtFile(sBatFilePath, asRange)
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sCMDEXEBAT_REDIRECT_FILE_NAME
    
    '*** コマンド実行 ***
    Open sOutputFilePath For Append As #1
    Print #1, ""
    Print #1, "****************************************************"
    Print #1, Now()
    Print #1, "****************************************************"
    Close #1
    Call ExecDosCmd(sBatFilePath & " >> " & sOutputFilePath, False)
    
    '*** バッチファイル削除 ***
    Kill sBatFilePath
    
    MsgBox "実行完了！", vbOKOnly, sMACRO_NAME
    
    '*** 出力ファイルを開く ***
'    If Left(sOutputFilePath, 1) = "" Then
'        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
'    Else
'        'Do Nothing
'    End If
'    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = 概要    選択範囲内のDOSコマンドをバッチファイルに書き出してまとめて実行する。（管理者権限）
' =         単一列選択時のみ有効。
' = 覚書    なし
' = 依存    Mng_SysCmd.bas/ExecDosCmdRunas()
' =         SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub DOSコマンドを一括実行_管理者権限()
    Const sMACRO_NAME As String = "DOSコマンドを一括実行_管理者権限"
    
    '*** アドイン設定読み出し ***
    Dim bIgnoreInvisibleCell As Boolean
    bIgnoreInvisibleCell = ReadSettingFile("bCMDEXEBATRUNAS_IGNORE_INVISIBLE_CELL", bCMDEXEBATRUNAS_IGNORE_INVISIBLE_CELL)
    
    '*** セル選択判定 ***
    If Selection.Count = 0 Then
        MsgBox "セルが選択されていません", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        End
    End If
    
    '*** 範囲チェック ***
    If Selection.Columns.Count = 1 Then
        'Do Nothing
    Else
        MsgBox "単一列のみ選択してください", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        End
    End If
    
    'Range型からString()型へ変換
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIgnoreInvisibleCell, _
                "" _
            )
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sCMDEXEBATRUNAS_REDIRECT_FILE_NAME
    
    '*** コマンド実行 ***
    Open sOutputFilePath For Append As #1
    Print #1, ""
    Print #1, "****************************************************"
    Print #1, Now()
    Print #1, "****************************************************"
    Print #1, ExecDosCmdRunas(asRange, True)
    Close #1
    
    MsgBox "実行完了！", vbOKOnly, sMACRO_NAME
    
    '*** 出力ファイルを開く ***
'    If Left(sOutputFilePath, 1) = "" Then
'        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
'    Else
'        'Do Nothing
'    End If
'    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = 概要    選択範囲内のDOSコマンドをそれぞれ実行する。
' =         単一列選択時のみ有効。
' = 覚書    ・単発のコマンドを実行する場合、「DOSコマンドを一括実行()」に比べて
' =           本マクロのほうが早い。
' =         ・大量のコマンドを実行する際、コマンド毎にプロンプトが表示される。
' =           目障りに感じる場合は、「DOSコマンドを一括実行()」を実行すること。
' = 依存    Mng_Array.bas/ConvRange2Array()
' =         Mng_SysCmd.bas/ExecDosCmd()
' =         SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub DOSコマンドを各々実行()
    Const sMACRO_NAME As String = "DOSコマンドを各々実行"
    
    '*** アドイン設定読み出し ***
    Dim bIgnoreInvisibleCell As Boolean
    bIgnoreInvisibleCell = ReadSettingFile("bCMDEXEUNI_IGNORE_INVISIBLE_CELL", bCMDEXEUNI_IGNORE_INVISIBLE_CELL)
    
    '*** セル選択判定 ***
    If Selection.Count = 0 Then
        MsgBox "セルが選択されていません", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        End
    End If
    
    '*** 範囲チェック ***
    If Selection.Columns.Count = 1 Then
        'Do Nothing
    Else
        MsgBox "単一列のみ選択してください", vbCritical, sMACRO_NAME
        MsgBox "処理を中断します", vbCritical, sMACRO_NAME
        End
    End If
    
    'Range型からString()型へ変換
    Dim asRange() As String
    Call ConvRange2Array( _
                Selection, _
                asRange, _
                bIgnoreInvisibleCell, _
                "" _
            )
    
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    Dim sOutputFilePath As String
    sOutputFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sCMDEXEUNI_REDIRECT_FILE_NAME
    
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
    
    MsgBox "実行完了！", vbOKOnly, sMACRO_NAME
    
    '*** 出力ファイルを開く ***
'    If Left(sOutputFilePath, 1) = "" Then
'        sOutputFilePath = Mid(sOutputFilePath, 2, Len(sOutputFilePath) - 2)
'    Else
'        'Do Nothing
'    End If
'    objWshShell.Run """" & sOutputFilePath & """", 3
End Sub

' =============================================================================
' = 概要    選択範囲内の検索文字の文字色を変更する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 検索文字の文字色を変更()
    Const sMACRO_NAME As String = "検索文字の文字色を変更"
    Const lSELECT_CLR_PALETTE As Boolean = True
    Const lREGEXP_IGNORECASE As Boolean = False
    
    Dim cCLR_RGBS As Variant
    Set cCLR_RGBS = CreateObject("System.Collections.ArrayList")
    '▼▼▼色設定▼▼▼
    Const sCOLOR_TYPE As String = "0:赤、1:水、2:緑、3:紫、4:橙、5:黄、6:白、7:黒"
    cCLR_RGBS.Add &HFF
    cCLR_RGBS.Add &HC6AC4B
    cCLR_RGBS.Add &H3C9376
    cCLR_RGBS.Add &HA03070
    cCLR_RGBS.Add &H4696F7
    cCLR_RGBS.Add &HC0FF
    cCLR_RGBS.Add &HFFFFFF
    cCLR_RGBS.Add &H0
    '▲▲▲色設定▲▲▲
    
    '*** アドイン設定ファイルから設定読み出し ***
    Dim sSrchStr As String
    Dim lClrRgbInit As Long
    sSrchStr = ReadSettingFile("sWORDCOLOR_SRCH_WORD", sWORDCOLOR_SRCH_WORD)
    lClrRgbInit = ReadSettingFile("lWORDCOLOR_CLR_RGB", lWORDCOLOR_CLR_RGB)
    
    '検索対象文字列選択
    sSrchStr = InputBox("検索文字列を正規表現で入力してください", sMACRO_NAME, sSrchStr)
    If StrPtr(sSrchStr) = 0 Then
        MsgBox "キャンセルが押されたため、処理を中断します。", vbCritical, sMACRO_NAME
        Exit Sub
    ElseIf sSrchStr = "" Then
        MsgBox "文字列が指定されなかったため、処理を中断します。", vbCritical, sMACRO_NAME
        Exit Sub
    Else
        'Do Nothing
    End If
    
    '色選択
    Dim lClrRgbSelected As Long
    If lSELECT_CLR_PALETTE = True Then 'カラーパレットで選択
        Dim bRet As Boolean
        bRet = ShowColorPalette(lClrRgbInit, lClrRgbSelected)
        If bRet = False Then
            MsgBox "色選択が失敗しましたので、処理を中断します。", vbCritical, sMACRO_NAME
            Exit Sub
        End If
    Else '色種別名で選択
        '色→色種別 変換
        Dim lClrTypeIdx As Long
        lClrTypeIdx = 0
        Dim bExist As Boolean
        bExist = False
        Dim vClrRgb As Variant
        For Each vClrRgb In cCLR_RGBS
            If vClrRgb = lClrRgbInit Then
                bExist = True
                Exit For
            Else
                lClrTypeIdx = lClrTypeIdx + 1
            End If
        Next
        If bExist = True Then
            'Do Nothing
        Else
            lClrTypeIdx = 0
        End If
        '色種別 選択
        lClrTypeIdx = InputBox( _
            "文字色を選択してください" & vbNewLine & _
            "  " & sCOLOR_TYPE & vbNewLine _
            , _
            sMACRO_NAME, _
            lClrTypeIdx _
        )
        '色種別→色 変換
        If lClrTypeIdx < cCLR_RGBS.Count Then
            lClrRgbSelected = cCLR_RGBS(lClrTypeIdx)
        Else
            MsgBox "文字色は指定の範囲内で選択してください。" & vbNewLine & sCOLOR_TYPE, vbOKOnly, sMACRO_NAME
            Exit Sub
        End If
    End If
    
    'アドイン設定更新
    Call WriteSettingFile("sWORDCOLOR_SRCH_WORD", sSrchStr)
    Call WriteSettingFile("lWORDCOLOR_CLR_RGB", lClrRgbSelected)
    
    '対象範囲特定(選択範囲と使用されている範囲の共通部分)
    Dim rTrgtRng As Range
    Set rTrgtRng = Application.Intersect(Selection, ActiveSheet.UsedRange)
    
    '検索文字列色変更
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSrchStr
    oRegExp.IgnoreCase = lREGEXP_IGNORECASE
    oRegExp.Global = True
    Dim oMatchResult
    Dim oCell As Range
    For Each oCell In rTrgtRng
        If oCell.Value <> "" Then
            Dim sTargetStr
            sTargetStr = oCell.Value
            Set oMatchResult = oRegExp.Execute(sTargetStr)
            Dim lMatchIdx As Long
            For lMatchIdx = 0 To oMatchResult.Count - 1
                Dim lCharPos As Long
                lCharPos = oMatchResult(lMatchIdx).FirstIndex + 1
                oCell.Characters( _
                    Start:=lCharPos, _
                    Length:=oMatchResult(lMatchIdx).Length _
                ).Font.Color = lClrRgbSelected
            Next lMatchIdx
        End If
    Next
    Set oMatchResult = Nothing
    Set oRegExp = Nothing
    
    MsgBox "完了！", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = 概要    ①～⑭を指定して、指定番号以降をデクリメントする
' = 覚書    なし
' = 依存    Mng_String.bas/NumConvStr2Lng()
' =         Mng_String.bas/NumConvLng2Str()
' = 所属    Macros.bas
' =============================================================================
Public Sub セル内の丸数字をデクリメント()
    Const sMACRO_NAME As String = "セル内の丸数字をデクリメント"
    Const NUM_MAX As Long = 15
    Const NUM_MIN As Long = 1
    
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
    MsgBox "置換完了！", vbOKOnly, sMACRO_NAME
End Sub

' =============================================================================
' = 概要    ②～⑮を指定して、指定番号以降をインクリメントする
' = 覚書    なし
' = 依存    Mng_String.bas/NumConvStr2Lng()
' =         Mng_String.bas/NumConvLng2Str()
' = 所属    Macros.bas
' =============================================================================
Public Sub セル内の丸数字をインクリメント()
    Const sMACRO_NAME As String = "セル内の丸数字をインクリメント"
    Const NUM_MAX As Long = 15
    Const NUM_MIN As Long = 1
    
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
    MsgBox "置換完了！", vbOKOnly, sMACRO_NAME
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
' = 概要    選択した範囲のハイパーリンクを一括で開く
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub ハイパーリンク一括オープン()
    Const sMACRO_NAME As String = "ハイパーリンク一括オープン"
    Dim Rng As Range
    
    If TypeName(Selection) = "Range" Then
        For Each Rng In Selection
            If Rng.Hyperlinks.Count > 0 Then Rng.Hyperlinks(1).Follow
        Next
    Else
        MsgBox "セル範囲が選択されていません。", vbExclamation, sMACRO_NAME
    End If
End Sub

' =============================================================================
' = 概要    アクティブセルからハイパーリンク先に飛ぶ
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub ハイパーリンクで飛ぶ()
    Dim rTrgtCell As Range
    On Error Resume Next
    For Each rTrgtCell In Selection
        rTrgtCell.Hyperlinks(1).Follow NewWindow:=True
        If Err.Number = 0 Then
            'Do Nothing
        Else
            Debug.Print "[" & Now & "] Error " & _
                        "[Macro] ハイパーリンクで飛ぶ " & _
                        "[Error No." & Err.Number & "] " & Err.Description
        End If
    Next
    On Error GoTo 0
End Sub

' =============================================================================
' = 概要    選択セルに対して「選択範囲内で中央」を実行する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 選択範囲内で中央()
    If Selection(1).HorizontalAlignment = xlCenterAcrossSelection Then
        Selection.HorizontalAlignment = xlGeneral
    Else
        Selection.HorizontalAlignment = xlCenterAcrossSelection
    End If
End Sub

' ==================================================================
' = 概要    選択範囲を範囲を維持したままセルコピーする。(ダブルクオーテーションを除く)
' = 覚書    ・セル内に改行が含まれる場合は範囲が崩れることに注意
' = 依存    Mng_Array.bas/ConvRange2Array()
' =         Mng_Clipboard.bas/SetToClipboard()
' =         SettingFile.cls
' = 所属    Macros.bas
' ==================================================================
Public Sub 範囲を維持したままセルコピー()
    Const sMACRO_NAME As String = "範囲を維持したままセルコピー"
    
    Application.ScreenUpdating = False
    
    '*** アドイン設定読み出し ***
    Dim bIgnoreInvisible As Boolean
    bIgnoreInvisible = ReadSettingFile("bCELLCOPYRNG_IGNORE_INVISIBLE_CELL", bCELLCOPYRNG_IGNORE_INVISIBLE_CELL)
    
    Dim sDelimiter As String
    sDelimiter = ReadSettingFile("sCELLCOPYRNG_DELIMITER", sCELLCOPYRNG_DELIMITER)
    
    '*** 選択範囲取得 ***
    Dim sClipedText As String
    sClipedText = ""
    Dim lAreaIdx As Long
    For lAreaIdx = 1 To Selection.Areas.Count
        '*** 追加テキスト取得 ***
        Dim asLine() As String
        Call ConvRange2Array( _
            Selection.Areas(lAreaIdx), _
            asLine, _
            bIgnoreInvisible, _
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
        
        If lAreaIdx = 1 Then
            sClipedText = sNewText
        Else
            sClipedText = sClipedText & vbNewLine & sNewText
        End If
    Next lAreaIdx
    
    '*** クリップボード設定 ***
    Call SetToClipboard(sClipedText)
    
    Application.ScreenUpdating = True
    
    '*** フィードバック ***
    Application.StatusBar = "■■■■■■■■ " & sMACRO_NAME & "完了！ ■■■■■■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' =============================================================================
' = 概要    選択範囲を一行にまとめてセルコピーする。
' = 覚書    ・セル内に改行が含まれる場合は一行にまとめられないことに注意
' = 依存    Mng_Clipboard.bas/SetToClipboard()
' =         SettingFile.cls
' = 所属    Macro.bas
' =============================================================================
Public Sub 一行にまとめてセルコピー()
    Const sMACRO_NAME As String = "一行にまとめてセルコピー"
    
    Application.ScreenUpdating = False
    
    '*** アドイン設定読み出し ***
    Dim bIgnoreInvisibleCell As Boolean
    Dim bIgnoreBlankCell As Boolean
    Dim sPreffix As String
    Dim sDelimiter As String
    Dim sSuffix As String
    bIgnoreInvisibleCell = ReadSettingFile("bCELLCOPYLINE_IGNORE_INVISIBLE_CELL", bCELLCOPYLINE_IGNORE_INVISIBLE_CELL)
    bIgnoreBlankCell = ReadSettingFile("bCELLCOPYLINE_IGNORE_BLANK_CELL", bCELLCOPYLINE_IGNORE_BLANK_CELL)
    sPreffix = ReadSettingFile("sCELLCOPYLINE_PREFFIX", sCELLCOPYLINE_PREFFIX)
    sDelimiter = ReadSettingFile("sCELLCOPYLINE_DELIMITER", sCELLCOPYLINE_DELIMITER)
    sSuffix = ReadSettingFile("sCELLCOPYLINE_SUFFIX", sCELLCOPYLINE_SUFFIX)
    
    '*** 選択範囲取得 ***
    Dim sClipedText As String
    sClipedText = ""
    Dim lAreaIdx As Long
    For lAreaIdx = 1 To Selection.Areas.Count
        Dim lItemIdx As Long
        For lItemIdx = 1 To Selection.Areas(lAreaIdx).Count
            With Selection.Areas(lAreaIdx).Item(lItemIdx)
                If .Value = "" Then                                     '空白セル
                    If bIgnoreBlankCell = True Then
                        'Do Nothing
                    Else
                        If sClipedText = "" Then
                            sClipedText = sPreffix & .Value
                        Else
                            sClipedText = sClipedText & sDelimiter & .Value
                        End If
                    End If
                Else
                    If .EntireRow.Hidden Or .EntireColumn.Hidden Then   '非表示セル
                        If bIgnoreInvisibleCell = True Then
                            'Do Nothing
                        Else
                            If sClipedText = "" Then
                                sClipedText = sPreffix & .Value
                            Else
                                sClipedText = sClipedText & sDelimiter & .Value
                            End If
                        End If
                    Else                                                '上記以外
                        If sClipedText = "" Then
                            sClipedText = sPreffix & .Value
                        Else
                            sClipedText = sClipedText & sDelimiter & .Value
                        End If
                    End If
                End If
            End With
        Next lItemIdx
    Next lAreaIdx
    sClipedText = sClipedText & sSuffix
    
    '*** クリップボード設定 ***
    Call SetToClipboard(sClipedText)
    
    Application.ScreenUpdating = True
    
    '*** フィードバック ***
    Application.StatusBar = "■■■■■■■■ " & sMACRO_NAME & "完了！ ■■■■■■■■"
    Sleep 200 'ms 単位
    Application.StatusBar = False
End Sub

' =============================================================================
' = 概要    一行にまとめてセルコピーにて使用する「先頭文字,区切り文字,末尾文字」を変更する
' = 覚書    なし
' = 依存    SettingFile.cls
' = 所属    Macro.bas
' =============================================================================
Public Sub ●設定変更●一行にまとめてセルコピー()
    Const sMACRO_NAME As String = "●設定変更●一行にまとめてセルコピー"
    
    Application.ScreenUpdating = False
    
    '*** アドイン設定読み出し ***
    Dim sPreffix As String
    Dim sDelimiter As String
    Dim sSuffix As String
    sPreffix = ReadSettingFile("sCELLCOPYLINE_PREFFIX", sCELLCOPYLINE_PREFFIX)
    sDelimiter = ReadSettingFile("sCELLCOPYLINE_DELIMITER", sCELLCOPYLINE_DELIMITER)
    sSuffix = ReadSettingFile("sCELLCOPYLINE_SUFFIX", sCELLCOPYLINE_SUFFIX)
    
    Dim vRet As Variant
    vRet = MsgBox( _
        "「" & sMACRO_NAME & "」の設定を変更します。" & vbNewLine & _
        "　先頭文字：" & sPreffix & vbNewLine & _
        "　区切り文字：" & sDelimiter & vbNewLine & _
        "　末尾文字：" & sSuffix & vbNewLine & _
        "" & vbNewLine & _
        "新たに設定を変更しますか？(→はい)" & vbNewLine & _
        "デフォルトの設定に戻しますか？(→いいえ)", _
        vbYesNoCancel, _
        sMACRO_NAME _
    )
    If vRet = vbYes Then
        sPreffix = InputBox( _
            "「先頭文字」を指定してください", _
            sMACRO_NAME, _
            sPreffix _
        )
        sDelimiter = InputBox( _
            "「区切り文字」を指定してください", _
            sMACRO_NAME, _
            sDelimiter _
        )
        sSuffix = InputBox( _
            "「末尾文字」を指定してください", _
            sMACRO_NAME, _
            sSuffix _
        )
        Call WriteSettingFile("sCELLCOPYLINE_PREFFIX", sPreffix)
        Call WriteSettingFile("sCELLCOPYLINE_DELIMITER", sDelimiter)
        Call WriteSettingFile("sCELLCOPYLINE_SUFFIX", sSuffix)
        MsgBox _
            "設定を変更しました" & vbNewLine & _
            "　先頭文字：" & sPreffix & vbNewLine & _
            "　区切り文字：" & sDelimiter & vbNewLine & _
            "　末尾文字：" & sSuffix, _
            vbOKOnly, _
            sMACRO_NAME
    ElseIf vRet = vbNo Then
        Call WriteSettingFile("sCELLCOPYLINE_PREFFIX", sPreffix)
        Call WriteSettingFile("sCELLCOPYLINE_DELIMITER", sDelimiter)
        Call WriteSettingFile("sCELLCOPYLINE_SUFFIX", sSuffix)
        Application.ScreenUpdating = True
        MsgBox _
            "設定をデフォルトに戻しました" & vbNewLine & _
            "　先頭文字：" & sPreffix & vbNewLine & _
            "　区切り文字：" & sDelimiter & vbNewLine & _
            "　末尾文字：" & sSuffix, _
            vbOKOnly, _
            sMACRO_NAME
    Else
        Application.ScreenUpdating = True
        MsgBox "処理をキャンセルします", vbExclamation, sMACRO_NAME
    End If
End Sub

' =============================================================================
' = 概要    クリップボードから値貼り付けする
' = 覚書    ・現在の選択範囲は無視する
' = 依存    Mng_Clipboard.bas/GetFromClipboard()
' = 所属    Macro.bas
' =============================================================================
Public Sub クリップボード値貼り付け()
    Dim bResult As Boolean
    Dim sStr As String
    bResult = GetFromClipboard(sStr)
    If bResult = True Then
        ActiveSheet.PasteSpecial Format:="テキスト"
    Else
        'Do Nothing
    End If
End Sub

' =============================================================================
' = 概要    フォント色を「設定色」⇔「自動」でトグルする
' = 覚書    なし
' = 依存    SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub フォント色をトグル()
    'アドイン設定読み出し
    Dim lClrRgb As Long
    lClrRgb = ReadSettingFile("lCLRTGLFONT_CLR_RGB", lCLRTGLFONT_CLR_RGB)
    
    'フォント色変更
    If Selection(1).Font.Color = lClrRgb Then
        Selection.Font.ColorIndex = xlAutomatic
    Else
        Selection.Font.Color = lClrRgb
    End If
End Sub

' =============================================================================
' = 概要    「フォント色をトグル」の設定色をカラーパレットから取得して変更する
' = 覚書    なし
' = 依存    SettingFile.cls
' =         Macros.bas/ShowColorPalette()
' = 所属    Macros.bas
' =============================================================================
Public Sub ●設定変更●フォント色をトグルの色選択()
    Const sMACRO_NAME As String = "●設定変更●フォント色をトグルの色選択"
    
    MsgBox sMACRO_NAME & "を実行します", vbOKOnly, sMACRO_NAME
    
    'アドイン設定読み出し
    Dim lClrRgbInit As Long
    lClrRgbInit = ReadSettingFile("lCLRTGLFONT_CLR_RGB", lCLRTGLFONT_CLR_RGB)
    
    '色選択
    Dim bRet As Boolean
    Dim lClrRgbSelected As Long
    bRet = ShowColorPalette(lClrRgbInit, lClrRgbSelected)
    If bRet = False Then
        MsgBox "色選択が失敗しましたので、処理を中断します。", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    
    'アドイン設定更新
    Call WriteSettingFile("lCLRTGLFONT_CLR_RGB", lClrRgbSelected)
End Sub

' =============================================================================
' = 概要    「フォント色をトグル」の設定色をアクティブセルから取得して変更する
' = 覚書    なし
' = 依存    SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub ●設定変更●フォント色をトグルの色スポイト()
    Const sMACRO_NAME As String = "●設定変更●フォント色をトグルの色スポイト"
    
    '色取得
    Dim lClrRgb As Long
    lClrRgb = Selection(1).Font.Color
    
    'アドイン設定更新
    Call WriteSettingFile("lCLRTGLFONT_CLR_RGB", lClrRgb)
End Sub

' =============================================================================
' = 概要    背景色を「設定色」⇔「背景色なし」でトグルする
' = 覚書    なし
' = 依存    SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub 背景色をトグル()
    'アドイン設定読み出し
    Dim lClrRgb As Long
    lClrRgb = ReadSettingFile("lCLRTGLBG_CLR_RGB", lCLRTGLBG_CLR_RGB)
    
    '背景色変更
    If Selection(1).Interior.Color = lClrRgb Then
        Selection.Interior.ColorIndex = 0
    Else
        Selection.Interior.Color = lClrRgb
    End If
End Sub

' =============================================================================
' = 概要    「背景色をトグル」の設定色をカラーパレットから取得して変更する
' = 覚書    なし
' = 依存    SettingFile.cls
' =         Macros.bas/ShowColorPalette()
' = 所属    Macros.bas
' =============================================================================
Public Sub ●設定変更●背景色をトグルの色選択()
    Const sMACRO_NAME As String = "●設定変更●背景色をトグルの色選択"
    
    MsgBox sMACRO_NAME & "を実行します", vbOKOnly, sMACRO_NAME
    
    'アドイン設定読み出し
    Dim lClrRgbInit As Long
    lClrRgbInit = ReadSettingFile("lCLRTGLBG_CLR_RGB", lCLRTGLBG_CLR_RGB)
    
    '色選択
    Dim bRet As Boolean
    Dim lClrRgbSelected As Long
    bRet = ShowColorPalette(lClrRgbInit, lClrRgbSelected)
    If bRet = False Then
        MsgBox "色選択が失敗しましたので、処理を中断します。", vbCritical, sMACRO_NAME
        Exit Sub
    End If
    
    'アドイン設定更新
    Call WriteSettingFile("lCLRTGLBG_CLR_RGB", lClrRgbSelected)
End Sub

' =============================================================================
' = 概要    「背景色をトグル」の設定色をアクティブセルから取得して変更する
' = 覚書    なし
' = 依存    SettingFile.cls
' = 所属    Macros.bas
' =============================================================================
Public Sub ●設定変更●背景色をトグルの色スポイト()
    Const sMACRO_NAME As String = "●設定変更●背景色をトグルの色スポイト"
    
    '色取得
    Dim lClrRgb As Long
    lClrRgb = Selection(1).Interior.Color
    
    'アドイン設定更新
    Call WriteSettingFile("lCLRTGLBG_CLR_RGB", lClrRgb)
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
' = 概要    画面を上に移動(スクロールロック動作)
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 画面を上に移動()
    With ActiveWindow
        If .ScrollRow > 1 Then
            .ScrollRow = .ScrollRow - 1
        End If
    End With
End Sub

' =============================================================================
' = 概要    画面を下に移動(スクロールロック動作)
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 画面を下に移動()
    With ActiveWindow
        .ScrollRow = .ScrollRow + 1
    End With
End Sub

' =============================================================================
' = 概要    画面を左に移動(スクロールロック動作)
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 画面を左に移動()
    With ActiveWindow
        If .ScrollColumn > 1 Then
            .ScrollColumn = .ScrollColumn - 1
        End If
    End With
End Sub

' =============================================================================
' = 概要    画面を右に移動(スクロールロック動作)
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub 画面を右に移動()
    With ActiveWindow
        .ScrollColumn = .ScrollColumn + 1
    End With
End Sub

' =============================================================================
' = 概要    インデントを上げる
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub インデントを上げる()
    Dim rCell As Range
    For Each rCell In Selection
        rCell.IndentLevel = rCell.IndentLevel + 1
    Next
End Sub

' =============================================================================
' = 概要    インデントを下げる
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub インデントを下げる()
    Dim rCell As Range
    For Each rCell In Selection
        If rCell.IndentLevel = 0 Then
            'Do Nothing
        Else
            rCell.IndentLevel = rCell.IndentLevel - 1
        End If
    Next
End Sub

' =============================================================================
' = 概要    他セルコメントを“非表示”にしてアクティブセルコメントを“表示”(+移動)
' = 覚書    なし
' = 依存    Macros.bas/VisibleCommentOnlyActiveCell()
' = 所属    Macros.bas
' =============================================================================
Public Sub アクティブセルコメントのみ表示()
'   Application.ScreenUpdating = False
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub
Public Sub アクティブセルコメントのみ表示して下移動()
'   Application.ScreenUpdating = False
    ActiveCell.Offset(1, 0).Activate
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub
Public Sub アクティブセルコメントのみ表示して上移動()
'   Application.ScreenUpdating = False
    ActiveCell.Offset(-1, 0).Activate
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub
Public Sub アクティブセルコメントのみ表示して左移動()
'   Application.ScreenUpdating = False
    ActiveCell.Offset(0, -1).Activate
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub
Public Sub アクティブセルコメントのみ表示して右移動()
'   Application.ScreenUpdating = False
    ActiveCell.Offset(0, 1).Activate
    Call VisibleCommentOnlyActiveCell
'   Application.ScreenUpdating = True
End Sub

' =============================================================================
' = 概要    アクティブセルのコメント表示の有効/無効を切り替える
' = 覚書    なし
' = 依存    SettingFile.cls
' =         Macros.bas/SwitchMacroShortcutKeysActivation()
' = 所属    Macros.bas
' =============================================================================
Public Sub ●設定変更●アクティブセルコメントのみ表示()
    Const sMACRO_NAME As String = "●設定変更●アクティブセルコメントのみ表示"
    
    'アドイン設定ファイル読み出し
    Dim bExistSetting As Boolean
    bExistSetting = ReadSettingFile("bCMNT_VSBL_ENB", bCMNT_VSBL_ENB)
    
    'アクティブセルコメント設定更新
    Dim bCmntVsblEnb As Boolean
    If bExistSetting = True Then
        If bCmntVsblEnb = True Then
            MsgBox "アクティブセルコメントのみ表示を【無効化】します", vbOKOnly, sMACRO_NAME
            bCmntVsblEnb = False
        Else
            MsgBox "アクティブセルコメントのみ表示を【有効化】します", vbOKOnly, sMACRO_NAME
            bCmntVsblEnb = True
        End If
    Else
        MsgBox "アクティブセルコメントのみ表示を【有効化】します", vbOKOnly, sMACRO_NAME
        bCmntVsblEnb = True
    End If
    
    Call WriteSettingFile("bCMNT_VSBL_ENB", bCmntVsblEnb)
    
    'ショートカットキー設定 更新(有効化)
    Call SwitchMacroShortcutKeysActivation(True)
End Sub

' =============================================================================
' = 概要    Excel数式整形化実施/解除
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub Excel数式整形化実施()
    Dim rSelectRange As Range
    For Each rSelectRange In Selection
        rSelectRange.Formula = ConvFormuraIndentation(rSelectRange.Formula, True)
    Next
End Sub
Public Sub Excel数式整形化解除()
    Dim rSelectRange As Range
    For Each rSelectRange In Selection
        rSelectRange.Formula = ConvFormuraIndentation(rSelectRange.Formula, False)
    Next
End Sub

' =============================================================================
' = 概要    現在シートの全セルコメントオブジェクトを
' =         「セルに合わせて移動やサイズ変更をする」に一括設定する
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub セルコメントの書式設定を一括変更()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    With ActiveSheet
        Dim lLastRow As Long
        Dim lLastClm As Long
        Dim lRowIdx As Long
        Dim lClmIdx As Long
        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lLastClm = .Cells(1, .Columns.Count).End(xlToLeft).Column
        For lRowIdx = 1 To lLastRow
            For lClmIdx = 1 To lLastClm
                If .Cells(lRowIdx, lClmIdx).Comment Is Nothing Then
                    'Do Nothing
                Else
                    'MsgBox .Cells(lRowIdx, lClmIdx).Value
                    .Cells(lRowIdx, lClmIdx).Comment.Shape.Placement = xlMoveAndSize
                End If
            Next lClmIdx
        Next lRowIdx
    End With
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' = 概要    選択範囲のDiff形式のフォント色に変更する。(旧:赤、新:緑)
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub Diff色付け()
    Const bUNIFIED_MODE As Boolean = True
    Dim rCell As Range
    For Each rCell In Selection
        Dim oRegExp As Object
        Set oRegExp = CreateObject("VBScript.RegExp")
        Dim sTargetStr As String
        sTargetStr = rCell.Value
        oRegExp.IgnoreCase = True
        oRegExp.Global = True
        Dim oMatchResult As Object
        
        If bUNIFIED_MODE = True Then
            oRegExp.Pattern = "^\+"
        Else
            oRegExp.Pattern = "^>"
        End If
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        If oMatchResult.Count > 0 Then
            rCell.Font.Color = RGB(0, 176, 80)
        End If
        
        If bUNIFIED_MODE = True Then
            oRegExp.Pattern = "^-"
        Else
            oRegExp.Pattern = "^<"
        End If
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        If oMatchResult.Count > 0 Then
            rCell.Font.Color = RGB(255, 0, 0)
        End If
        
        oRegExp.Pattern = "^\$ diff"
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        If oMatchResult.Count > 0 Then
            rCell.Font.Bold = True
        End If
        
        oRegExp.Pattern = "^\$ git diff"
        Set oMatchResult = oRegExp.Execute(sTargetStr)
        If oMatchResult.Count > 0 Then
            rCell.Font.Bold = True
        End If
    Next
End Sub

' ▽▽▽ オブジェクト操作 ▽▽▽
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
' = 概要    現在シートのを全オブジェクトを
' =         「セルに合わせて移動とサイズ変更をする」に変更
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Public Sub オブジェクトサイズ変更プロパティ一括変更()
    Dim objShp As Shape
    For Each objShp In ActiveSheet.Shapes
        objShp.Placement = xlMoveAndSize
    Next
    MsgBox "アクティブシート内の全オブジェクトのプロパティを、" & vbNewLine & _
        "「セルに合わせて移動とサイズ変更をする」に一括変更しました！"
End Sub

' =============================================================================
' = 概要    選択範囲のセルアドレスを結合して文字列コピー
' = 覚書    なし
' = 依存    Macros.bas/CopyConcatedCellAddresses()
' = 所属    Macros.bas
' =============================================================================
Public Sub 選択範囲アドレス結合文字列コピー_相対行_相対列()
    Call CopyConcatedCellAddresses(Selection, False, False, bCELLADRJOIN_FORMAT_R1C1, sCELLADRJOIN_DELIMITER)
End Sub
Public Sub 選択範囲アドレス結合文字列コピー_絶対行_相対列()
    Call CopyConcatedCellAddresses(Selection, True, False, bCELLADRJOIN_FORMAT_R1C1, sCELLADRJOIN_DELIMITER)
End Sub
Public Sub 選択範囲アドレス結合文字列コピー_相対行_絶対列()
    Call CopyConcatedCellAddresses(Selection, False, True, bCELLADRJOIN_FORMAT_R1C1, sCELLADRJOIN_DELIMITER)
End Sub
Public Sub 選択範囲アドレス結合文字列コピー_絶対行_絶対列()
    Call CopyConcatedCellAddresses(Selection, True, True, bCELLADRJOIN_FORMAT_R1C1, sCELLADRJOIN_DELIMITER)
End Sub

' *****************************************************************************
' * 内部プロシージャ定義
' *****************************************************************************
Private Sub ▼▼▼▼▼内部プロシージャ▼▼▼▼▼()
    'プロシージャリスト表示用のダミープロシージャ
End Sub

' =============================================================================
' = 概要    シートを並び替える。
' =         シート並べ替え作業用シートに記載の通り、シートを並び替える。
' =         必ずシート並べ替え作業用シートから呼び出すこと！
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Private Sub SortSheetPost()
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
' = 概要    他セルコメントを“非表示”にしてアクティブセルコメントを“表示”にする。
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' =============================================================================
Private Sub VisibleCommentOnlyActiveCell()
    On Error Resume Next
    
    '全セルコメント非表示
    Dim cmComment As Comment
    For Each cmComment In ActiveSheet.Comments
        cmComment.Visible = False
    Next cmComment
    
    'アクティブセルコメント表示
    ActiveCell.Comment.Visible = True
    
    On Error GoTo 0
End Sub

' ==================================================================
' = 概要    アドイン設定用のファイルパスを取得する
' = 引数    なし
' = 戻値                    String      アドイン設定用のファイルパス
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Public Function GetAddinSettingFilePath() As String
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    GetAddinSettingFilePath = GetAddinSettingDirPath() & "\" & objFSO.GetBaseName(ThisWorkbook.Name) & ".cfg"
End Function

' ==================================================================
' = 概要    アドイン設定用のフォルダパスを取得する
' = 引数    なし
' = 戻値                    String      アドイン設定用のフォルダパス
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Public Function GetAddinSettingDirPath() As String
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    GetAddinSettingDirPath = _
        objWshShell.SpecialFolders("MyDocuments") & "\" & objFSO.GetBaseName(ThisWorkbook.Name)
End Function

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
' = 引数    bIgnoreInvisibleCell    String  [in]  非表示セル無視実行可否
' = 引数    sDelimiter              String  [in]  区切り文字
' = 戻値    なし
' = 覚書    列が隣り合ったセル同士は指定された区切り文字で区切られる
' = 依存    なし
' = 所属    Mng_Array.bas
' ==================================================================
Private Function ConvRange2Array( _
    ByRef rCellsRange As Range, _
    ByRef asLine() As String, _
    ByVal bIgnoreInvisibleCell As Boolean, _
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
            If bIgnoreInvisibleCell = True Then
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
' = 概要    コマンドを実行
' = 引数    sCommand    String   [in]   コマンド
' = 引数    bGetStdout  Boolean  [in]   標準出力取得有無(省略可)
' = 戻値                String          標準出力
' = 覚書    ・大量の処理を行うbatを実行する場合、bGetStdoutをFalseにすること。
' =           コマンドの実行結果が必要な場合は、コマンドにリダイレクトを含めること。
' =             例）Call ExecDosCmd("xxx.bat > xxx.log", False)
' =           【理由】
' =           Execは標準出力にためるバッファの最大は4096バイトであり、
' =           それ以上のデータを読み込むとAtEndOfStream時に固まるため。
' =           https://community.cybozu.dev/t/topic/181/2
' = 依存    なし
' = 所属    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmd( _
    ByVal sCommand As String, _
    Optional bGetStdOut As Boolean = True _
) As String
    If sCommand = "" Then
        ExecDosCmd = ""
    Else
        Dim sStdOutAll As String
        sStdOutAll = ""
        If bGetStdOut = True Then
            Dim oExeResult As Object
            Set oExeResult = CreateObject("WScript.Shell").Exec("%ComSpec% /c """ & sCommand & """")
            Do While Not oExeResult.StdOut.AtEndOfStream
                Dim sStdOut As String
                sStdOut = oExeResult.StdOut.ReadLine
                Debug.Print sStdOut
                sStdOutAll = sStdOutAll & vbNewLine & sStdOut
            Loop
            Set oExeResult = Nothing
        Else
            Call CreateObject("WScript.Shell").Run("%ComSpec% /c """ & sCommand & """", WaitOnReturn:=True)
        End If
        ExecDosCmd = sStdOutAll
    End If
End Function

' ==================================================================
' = 概要    コマンドを実行（管理者権限）
' = 引数    asCommands()    String   [in] 実行コマンド
' = 引数    bDelFiles       Boolean  [in] Bat/Logファイル削除(省略可)
' = 戻値                    String        標準出力＆標準エラー出力
' = 覚書    ・Desktopフォルダパスに空白が含まれる場合は、動作しない。
' = 依存    なし
' = 依存    Mng_FileSys.bas/OutputTxtFile()
' = 所属    Mng_SysCmd.bas
' ==================================================================
Private Function ExecDosCmdRunas( _
    ByRef asCommands() As String, _
    Optional bDelFiles As Boolean = True _
) As String
    Const sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME As String = "CmdExeBatRunas"
    If Sgn(asCommands) = 0 Then
        ExecDosCmdRunas = ""
    Else
        If UBound(asCommands) < 0 Then
            ExecDosCmdRunas = ""
        Else
            Dim objWshShell
            Set objWshShell = CreateObject("WScript.Shell")
            Dim objFSO
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            
            Dim sBatFilePath As String
            Dim sLogFilePath As String
            sBatFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME & ".bat"
            sLogFilePath = objWshShell.SpecialFolders("Desktop") & "\" & sEXECDOSCMDRUNAS_REDIRECT_FILE_BASE_NAME & ".log"
            
            '「@echo off」挿入
            ReDim Preserve asCommands(UBound(asCommands) + 1)
            Dim lIdx As Long
            For lIdx = UBound(asCommands) To (LBound(asCommands) + 1) Step -1
                asCommands(lIdx) = asCommands(lIdx - 1)
            Next lIdx
            asCommands(0) = "@echo off"
            
            'BATファイル作成
            Call OutputTxtFile(sBatFilePath, asCommands)
            Do While Not objFSO.FileExists(sBatFilePath)
                Sleep 100
            Loop
            
            'BATファイル実行
            ShellExecute 0, "runas", sBatFilePath, " > " & sLogFilePath & " 2>&1", vbNullString, 1
            
            'LOGファイル出力待ち
            Do While Not objFSO.FileExists(sLogFilePath)
                Sleep 100
            Loop
            
            'LOGファイル読込み
            Dim objTxtFile
            Set objTxtFile = objFSO.OpenTextFile(sLogFilePath, 1, True)
            Dim sStdOutAll As String
            sStdOutAll = ""
            Dim sLine As String
            Do Until objTxtFile.AtEndOfStream
                sLine = objTxtFile.ReadLine
                'MsgBox sLine
                If sStdOutAll = "" Then
                    sStdOutAll = sLine
                Else
                    sStdOutAll = sStdOutAll & vbNewLine & sLine
                End If
            Loop
            'MsgBox sStdOutAll
            objTxtFile.Close
            
            'BATファイル/LOGファイル削除
            If bDelFiles = True Then
                Kill sBatFilePath
                Kill sLogFilePath
            End If
            
            ExecDosCmdRunas = sStdOutAll
        End If
    End If
End Function
    Private Sub Test_ExecDosCmdRunas()
        Dim asCommands() As String
        
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(0)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target.txt"""
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(1)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target.txt"""
        asCommands(1) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source2.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target2.txt"""
        MsgBox ExecDosCmdRunas(asCommands)
        
        ReDim Preserve asCommands(1)
        asCommands(0) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target.txt"""
        asCommands(1) = "mklink ""C:\Users\draem\OneDrive\デスクトップ\source2.txt"" ""C:\Users\draem\OneDrive\デスクトップ\target2.txt"""
        MsgBox ExecDosCmdRunas(asCommands, False)
    End Sub

' ============================================
' = 概要    配列の内容をファイルに書き込む。
' = 引数    sFilePath       String  [in]  出力するファイルパス
' =         asFileLine()    String  [in]  出力するファイルの内容
' =         sCharSet        String  [in]  文字コード(省略可)
' =                                         (UTF-8|UTF-16|Shift_JIS|EUC-JP|ISO-2022-JP|...)
' =         lLineSeparator  Long    [in]  改行コード(省略可)
' =                                         13:CR 10:LF -1:CRLF
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_Array.bas
' ============================================
Private Function OutputTxtFile( _
    ByVal sFilePath As String, _
    ByRef asFileLine() As String, _
    Optional ByVal sCharSet As String = "shift_jis", _
    Optional ByVal lLineSeparator As Long = -1 _
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
            .LineSeparator = lLineSeparator
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
' = 概要    フォルダ選択ダイアログを表示する
' = 引数    sInitPath       String  [in]  デフォルトフォルダパス（省略可）
' = 引数    sTitle          String  [in]  タイトル名（省略可）
' = 戻値                    String        選択フォルダ
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sTitle As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If sTitle = "" Then
        fdDialog.Title = "フォルダを選択してください（空欄の場合は親フォルダが選択されます）"
    Else
        fdDialog.Title = sTitle
    End If
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

' ==================================================================
' = 概要    ファイル（単一）選択ダイアログを表示する
' = 引数    sInitPath       String  [in]  デフォルトファイルパス（省略可）
' = 引数    sTitle          String  [in]  タイトル名（省略可）
' = 引数    sFilters        String  [in]  選択時のフィルタ（省略可）(※)
' = 戻値                    String        選択ファイル
' = 覚書    (※)ダイアログのフィルタ指定方法は以下。
' =              ex) 画像ファイル/*.gif; *.jpg; *.jpeg,テキストファイル/*.txt; *.csv
' =                    ・拡張子が複数ある場合は、";"で区切る
' =                    ・ファイル種別と拡張子は"/"で区切る
' =                    ・フィルタが複数ある場合、","で区切る
' =         sFilters が省略もしくは空文字の場合、フィルタをクリアする。
' = 依存    Mng_FileSys.bas/SetDialogFilters()
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function ShowFileSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sTitle As String = "", _
    Optional ByVal sFilters As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    If sTitle = "" Then
        fdDialog.Title = "ファイルを選択してください"
    Else
        fdDialog.Title = sTitle
    End If
    fdDialog.AllowMultiSelect = False
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
        ShowFileSelectDialog = ""
    Else
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
            ShowFileSelectDialog = sSelectedPath
        Else
            ShowFileSelectDialog = ""
        End If
    End If
    
    Set fdDialog = Nothing
End Function
 
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
        .Worksheets.Add(After:=.Worksheets(.Worksheets.Count)).Name = sSheetName
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

' ==================================================================
' = 概要    クリップボードにテキストを設定（Win32Apiを使用）
' = 引数    sInStr      String  [in]  設定対象文字列
' = 戻値                Boolean       設定結果
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
    ByVal sInStr As String _
) As Boolean
#If VBA7 Then
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hClipMemory As LongPtr
    Dim lX As LongPtr
#Else
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long
    Dim lX As Long
#End If
    Dim bResult As Boolean
    bResult = True
    
    hGlobalMemory = GlobalAlloc(GHND, LenB(sInStr) + 1)   '移動可能なグローバルメモリを割り当て
    lpGlobalMemory = GlobalLock(hGlobalMemory)          'ブロックをロックして、メモリへのfarポインタを取得
    lpGlobalMemory = lstrcpy(lpGlobalMemory, sInStr)      '文字列をグローバルメモリへコピー
    
    'メモリのロック解除
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "メモリのロックを解除できません" & vbCrLf & "処理が失敗しました"
        bResult = False
    Else
        'データをコピーするクリップボードを開く
        If OpenClipboard(0&) = 0 Then
            MsgBox "クリップボードを開くことができません" & vbCrLf & "処理が失敗しました"
            bResult = False
            Exit Function
        End If
        
        lX = EmptyClipboard()    'クリップボードの内容を消去
        hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory) 'データをクリップボードへコピー
    End If
    
    'クリップボードの状態チェック
    If CloseClipboard() = 0 Then
        MsgBox "クリップボードを閉じることができません"
        bResult = False
    End If
    SetToClipboard = bResult
End Function

' ==================================================================
' = 概要    クリップボードからテキストを取得（Win32Apiを使用）
' = 引数    sOutStr     String  [Out]   取得先文字列
' = 戻値                Boolean         取得結果
' = 覚書    Win32APIを使用する。
' =         ※ クリップボードは DataObject の PutInClipboard でも利用
' =            可能｡しかし､DataObject は参照設定が必要なうえ､特定のク
' =            リップボード形式には貼り付けされない｡（CF_UNICODETEXT
' =            のみで CF_TEXTへは貼り付けされない）
' =            上記のように DataObject を使用したくない場合に本関数
' =            を利用すること｡
' = 依存    user32/OpenClipboard()
' =         user32/CloseClipboard()
' =         user32/GetClipboardData()
' =         kernel32/GlobalLock()
' =         kernel32/GlobalUnlock()
' =         kernel32/lstrcpy()
' = 所属    Mng_Clipboard.bas
' ==================================================================
Public Function GetFromClipboard( _
    ByRef sOutStr As String _
) As Boolean
#If VBA7 Then
    Dim hClipMemory As LongPtr
    Dim lpClipMemory As LongPtr
#Else
    Dim hClipMemory As Long
    Dim lpClipMemory As Long
#End If
    Dim sStr As String
    Dim lRetVal As Long
    Dim bResult As Boolean
    bResult = True
    sOutStr = ""
    
    If OpenClipboard(0&) = 0 Then
        MsgBox "クリップボードを開くことができません" & vbCrLf & "処理が失敗しました"
        bResult = False
        Exit Function
    End If
    
    ' Obtain the handle to the global memory block that is referencing the text.
    hClipMemory = GetClipboardData(CF_TEXT)
    If IsNull(hClipMemory) Then
        MsgBox "Could not allocate memory"
        bResult = False
    Else
        ' Lock Clipboard memory so we can reference the actual data string.
        lpClipMemory = GlobalLock(hClipMemory)
        
        If Not IsNull(lpClipMemory) Then
            sStr = Space$(MAXSIZE)
            Call lstrcpy(sStr, lpClipMemory)
            Call GlobalUnlock(hClipMemory)
            sStr = Mid(sStr, 1, InStr(1, sStr, Chr$(0), 0) - 1)
        Else
            MsgBox "Could not lock memory to copy string from."
            bResult = False
        End If
    End If
    
    If CloseClipboard() = 0 Then
        MsgBox "クリップボードを閉じることができません"
        bResult = False
    Else
        sOutStr = sStr
    End If
    GetFromClipboard = bResult
End Function

' ==================================================================
' = 概要    絶対パスから検索キー配下階層の相対パスへ置換
' = 引数    sInFilePath     String  [in]    絶対パス
' = 引数    sMatchDirName   String  [in]    検索対象フォルダ名
' = 引数    lRemeveDirLevel Long    [in]    階層レベル
' = 引数    sRelativePath   String  [out]   相対パス
' = 戻値                    Boolean         検索結果
' = 覚書    実行例1)
'             sInFilePath     : c\codes\aaa\bbb\ccc\test.txt
'             sMatchDirName   : codes
'             lRemeveDirLevel : 1
'             ↓
'             sRelativePath   : bbb\ccc\test.txt
'             戻値            : true
'
'           実行例2)
'             sInFilePath     : c\codes\aaa\bbb\ccc\test.txt
'             sMatchDirName   : code
'             lRemeveDirLevel : 2
'             ↓
'             sRelativePath   : c\codes\aaa\bbb\ccc\test.txt
'             戻値            : false
' = 依存    なし
' = 所属    Mng_String.bas
' ==================================================================
Public Function ExtractRelativePath( _
    ByVal sInFilePath As String, _
    ByVal sMatchDirName As String, _
    ByVal lRemeveDirLevel As Long, _
    ByRef sRelativePath As String _
) As Boolean
    Dim sRemoveDirLevelPath
    sRemoveDirLevelPath = ""
    Dim lIdx
    For lIdx = 0 To lRemeveDirLevel - 1
        sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
    Next
    
    Dim sSearchPattern
    Dim sTargetStr
    sSearchPattern = ".*\\" & sMatchDirName & sRemoveDirLevelPath & "\\"
    sTargetStr = sInFilePath
    
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSearchPattern                '検索パターンを設定
    oRegExp.IgnoreCase = True                       '大文字と小文字を区別しない
    oRegExp.Global = True                           '文字列全体を検索
    
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)  'パターンマッチ実行
    
    If oMatchResult.Count > 0 Then
        sRelativePath = Replace(sInFilePath, oMatchResult.Item(0), "")
        ExtractRelativePath = True
    Else
        sRelativePath = sInFilePath
        ExtractRelativePath = False
    End If
End Function

' ==================================================================
' = 概要    末尾区切り文字以降の文字列を返却する。
' = 引数    sStr        String  [in]  分割する文字列
' = 引数    sDlmtr      String  [in]  区切り文字
' = 戻値                String        抽出文字列
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_String.bas
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
' = 概要    Excel数式を整形する
' = 引数    sInputCellFormula   String   [in]   入力数式
' = 引数    bExecIndentation    Boolean  [in]   整形実施/整形解除
' = 引数    lIndentWidth        Long     [in]   インデント文字数(省略可)
' = 戻値                        String          出力数式
' = 覚書    ・整形解除時は、数式に関係のない空白はすべて除去する
' = 依存    なし
' = 所属    Mng_ExcelOpe.bas
' ==================================================================
Private Function ConvFormuraIndentation( _
    ByVal sInputCellFormula As String, _
    ByVal bExecIndentation As Boolean, _
    Optional ByVal lIndentWidth As Long = 2 _
) As String
    Dim sOutputCellFormula As String
    sOutputCellFormula = ""
    
    '数式の場合
    If Left(sInputCellFormula, 1) = "=" Then
        Dim bStrMode As Boolean
        Dim lNestCnt As Long
        bStrMode = False
        lNestCnt = 0
        '文字列操作
        Dim lChrIdx As Long
        For lChrIdx = 1 To Len(sInputCellFormula)
            Dim sInputCellFormulaChr As String
            sInputCellFormulaChr = Mid(sInputCellFormula, lChrIdx, 1)
            
            '文字列モードの場合
            If bStrMode = True Then
                Select Case sInputCellFormulaChr
                Case """"
                    sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    bStrMode = False
                Case Else
                    sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                End Select
            '文字列モードでない場合
            Else
                Select Case sInputCellFormulaChr
                Case ","
                    If bExecIndentation = True Then
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr & vbLf & String(lNestCnt * lIndentWidth, " ")
                    Else
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    End If
                Case "("
                    If bExecIndentation = True Then
                        lNestCnt = lNestCnt + 1
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr & vbLf & String(lNestCnt * lIndentWidth, " ")
                    Else
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    End If
                Case ")"
                    If bExecIndentation = True Then
                        lNestCnt = lNestCnt - 1
                        sOutputCellFormula = sOutputCellFormula & vbLf & String(lNestCnt * lIndentWidth, " ") & sInputCellFormulaChr
                    Else
                        sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    End If
                Case """"
                    sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                    bStrMode = True
                Case vbLf
                    'Do Nothing
                Case " "
                    'Do Nothing
                Case Else
                    sOutputCellFormula = sOutputCellFormula & sInputCellFormulaChr
                End Select
            End If
        Next lChrIdx
    '数式でない場合
    Else
        sOutputCellFormula = sInputCellFormula
    End If
    
    ConvFormuraIndentation = sOutputCellFormula
End Function

' ==================================================================
' = 概要    色の設定ダイアログを表示し、そこで選択された色のRGB値を返す
' = 引数    lClrRgbInit       Long    [in]    RGB値 初期値
' = 引数    lClrRgbSelected   Long    [out]   RGB値 選択値
' = 戻値                      Boolean         選択結果
' =                                               (True:成功,False:キャンセルor失敗)
' = 覚書    ・キャンセルor失敗時、lClrRgbSelectedはInitと同じ値となる
' = 依存    なし
' = 所属    Mng_ExcelOpe.bas
' ==================================================================
Private Function ShowColorPalette( _
    ByVal lClrRgbInit As Long, _
    ByRef lClrRgbSelected As Long _
) As Boolean
    Const CC_RGBINIT = &H1          '色のデフォルト値を設定
    Const CC_LFULLOPEN = &H2        '色の作成を行う部分を表示
    Const CC_PREVENTFULLOPEN = &H4  '色の作成ボタンを無効にする
    Const CC_SHOWHELP = &H8         'ヘルプボタンを表示
    
    Dim tChooseColor As ChooseColor
    With tChooseColor
        'ダイアログの設定
        .lStructSize = Len(tChooseColor)
        .lpCustColors = String$(64, Chr$(0))
        .flags = CC_RGBINIT + CC_LFULLOPEN
        .rgbResult = lClrRgbInit
        
        'ダイアログを表示
        Dim lRet As Long
        lRet = ChooseColor(tChooseColor)
        
        'ダイアログからの返り値をチェック
        lClrRgbSelected = lClrRgbInit
        If lRet <> 0 Then
            If .rgbResult > RGB(255, 255, 255) Then 'エラー
                ShowColorPalette = False
            Else '正常終了
                ShowColorPalette = True
                lClrRgbSelected = .rgbResult
            End If
        Else 'キャンセル押下
            ShowColorPalette = False
        End If
    End With
End Function

' ==================================================================
' = 概要    設定ファイルから設定を取得する
' = 引数    sKey            String      [in]    設定キー
' = 引数    vInitValue      Variant     [in]    設定値(初期値)
' = 戻値                    Variant             設定値
' = 覚書    ・ファイルオープン後、設定値を取得する。
' =           設定値が存在しない場合、設定値(初期値)で設定値を更新して保存する。
' =         ・以下の場合、vInitValueを返却する
' =           - sFilePathが存在しない
' =           - sKeyが存在しない
' = 依存    Macros.bas/GetAddinSettingFilePath()
' =         Mng_FileSys.bas/CreateDirectry()
' =         Macros.bas/ConvCtrlchr2Str()
' =         Macros.bas/ConvStr2Ctrlchr()
' = 所属    Macros.bas
' ==================================================================
Public Function ReadSettingFile( _
    ByVal sKey As String, _
    ByVal vInitValue As Variant _
) As Variant
    Dim dSettingItems As Object
    Set dSettingItems = CreateObject("Scripting.Dictionary")
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim vKey As Variant
    
    '設定ファイルパス取得
    Dim sFilePath As String
    sFilePath = GetAddinSettingFilePath()
    
    '設定ファイル読み出し
    If objFSO.FileExists(sFilePath) Then
        'ファイル読み出し
        Open sFilePath For Input As #1
        Do Until EOF(1)
            Dim vKeyValue As Variant
            Dim sLine As String
            Line Input #1, sLine
            If InStr(sLine, sDELIMITER_INIT) Then
                vKeyValue = Split(sLine, sDELIMITER_INIT)
                If UBound(vKeyValue) = 0 Then
                    dSettingItems.Add vKeyValue(0), ""           '単一区切り文字(値なし)
                ElseIf UBound(vKeyValue) = 1 Then
                    dSettingItems.Add vKeyValue(0), vKeyValue(1) '単一区切り文字(値あり)
                Else
                    Stop                                          '複数区切り文字
                End If
            Else
                Stop                                              '区切り文字なし
            End If
        Loop
        Close #1
        
        '設定項目取得＆更新
        If dSettingItems.Exists(sKey) = True Then
            Dim sItem As String
            sItem = dSettingItems.Item(sKey)
            '型変換
            Dim vOutValue As Variant
            Select Case VarType(vInitValue)
                Case vbInteger: vOutValue = CInt(sItem)
                Case vbLong: vOutValue = CLng(sItem)
                Case vbSingle: vOutValue = CSng(sItem)
                Case vbDouble: vOutValue = CDbl(sItem)
                Case vbBoolean: vOutValue = CBool(sItem)
                Case vbString: vOutValue = CStr(ConvStr2Ctrlchr(sItem))
                Case vbCurrency: vOutValue = CCur(sItem)
                Case vbByte: vOutValue = CByte(sItem)
                Case vbDate: vOutValue = CDate(sItem)
                Case vbVariant: vOutValue = CVar(sItem)
               'Case vbEmpty      : vOutValue = CXxx(sItem)
               'Case vbNull       : vOutValue = CXxx(sItem)
               'Case vbObject     : vOutValue = CXxx(sItem)
               'Case vbError      : vOutValue = CXxx(sItem)
               'Case vbDataObject : vOutValue = CXxx(sItem)
               'Case vbArray      : vOutValue = CXxx(sItem)
                Case Else: vOutValue = ""
            End Select
            ReadSettingFile = vOutValue
        Else
            '項目追加
            dSettingItems.Add sKey, ConvCtrlchr2Str(CStr(vInitValue))
            
            'ファイル保存
            Open sFilePath For Output As #1
            For Each vKey In dSettingItems
                Print #1, vKey & sDELIMITER_INIT & dSettingItems.Item(vKey)
            Next
            Close #1
            ReadSettingFile = vInitValue
        End If
    Else
        '格納先フォルダ作成
        Call CreateDirectry(objFSO.GetParentFolderName(sFilePath))
        
        '項目追加
        dSettingItems.Add sKey, ConvCtrlchr2Str(CStr(vInitValue))
        
        'ファイル保存
        Open sFilePath For Output As #1
        For Each vKey In dSettingItems
            Print #1, vKey & sDELIMITER_INIT & dSettingItems.Item(vKey)
        Next
        Close #1
        
        ReadSettingFile = vInitValue
    End If
End Function

' ==================================================================
' = 概要    設定ファイルから設定を更新して保存する
' = 引数    sKey            String      [in]    設定キー
' = 引数    vValue          Variant     [in]    設定値
' = 戻値                                        なし
' = 覚書    ・ファイルオープン後、設定値を更新/追加する。
' = 依存    Macros.bas/GetAddinSettingFilePath()
' =         Mng_FileSys.bas/CreateDirectry()
' =         Macros.bas/ConvCtrlchr2Str()
' = 所属    Macros.bas
' ==================================================================
Public Function WriteSettingFile( _
    ByVal sKey As String, _
    ByVal vValue As Variant _
)
    Dim dSettingItems As Object
    Set dSettingItems = CreateObject("Scripting.Dictionary")
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim vKey As Variant
    
    '設定ファイルパス取得
    Dim sFilePath As String
    sFilePath = GetAddinSettingFilePath()
    
    '格納先フォルダ作成
    Call CreateDirectry(objFSO.GetParentFolderName(sFilePath))
    
    '設定ファイル読み出し
    Open sFilePath For Input As #1
    Do Until EOF(1)
        Dim vKeyValue As Variant
        Dim sLine As String
        Line Input #1, sLine
        If InStr(sLine, sDELIMITER_INIT) Then
            vKeyValue = Split(sLine, sDELIMITER_INIT)
            If UBound(vKeyValue) = 0 Then
                dSettingItems.Add vKeyValue(0), ""           '単一区切り文字(値なし)
            ElseIf UBound(vKeyValue) = 1 Then
                dSettingItems.Add vKeyValue(0), vKeyValue(1) '単一区切り文字(値あり)
            Else
                Stop                                          '複数区切り文字
            End If
        Else
            Stop                                              '区切り文字なし
        End If
    Loop
    Close #1
    
    '項目追加
    If dSettingItems.Exists(sKey) Then
        dSettingItems.Item(sKey) = vValue
    Else
        dSettingItems.Add sKey, ConvCtrlchr2Str(CStr(vValue))
    End If
    
    'ファイル保存
    Open sFilePath For Output As #1
    For Each vKey In dSettingItems
        Print #1, vKey & sDELIMITER_INIT & dSettingItems.Item(vKey)
    Next
    Close #1
End Function

' ==================================================================
' = 概要    設定値変換用 制御文字to文字列 変換
' = 引数    sValue          String      [in]    値(制御文字)
' = 戻値                    String              値(文字列)
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Private Function ConvCtrlchr2Str( _
    ByVal sValue As String _
) As String
    Select Case sValue
        Case vbTab:     ConvCtrlchr2Str = "vbTab"
        Case vbCr:      ConvCtrlchr2Str = "vbCr"
        Case vbLf:      ConvCtrlchr2Str = "vbLf"
        Case vbNewLine: ConvCtrlchr2Str = "vbNewLine"
        Case Else:      ConvCtrlchr2Str = sValue
    End Select
End Function

' ==================================================================
' = 概要    設定値変換用 文字列to制御文字 変換
' = 引数    sValue          String      [in]    値(文字列)
' = 戻値                    String              値(制御文字)
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Private Function ConvStr2Ctrlchr( _
    ByVal sValue As String _
) As String
    Select Case sValue
        Case "vbTab":     ConvStr2Ctrlchr = vbTab
        Case "vbCr":      ConvStr2Ctrlchr = vbCr
        Case "vbLf":      ConvStr2Ctrlchr = vbLf
        Case "vbNewLine": ConvStr2Ctrlchr = vbNewLine
        Case Else:        ConvStr2Ctrlchr = sValue
    End Select
End Function

' ==================================================================
' = 概要    ディレクトリを作成する。親ディレクトリも自動生成する。
' = 引数    sDirPath    String  [in]  フォルダパス
' = 戻値    なし
' = 覚書    フォルダが既に存在している場合は何もしない
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Private Function CreateDirectry( _
    ByVal sDirPath As String _
)
    Dim sParentDir As String
    Dim oFileSys As Object
    
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
    
    '親ディレクトリが存在しない場合、再帰呼び出し
    If oFileSys.FolderExists(sParentDir) = False Then
        Call CreateDirectry(sParentDir)
    End If
    
    'ディレクトリ作成
    If oFileSys.FolderExists(sDirPath) = False Then
        oFileSys.CreateFolder sDirPath
    End If
    
    Set oFileSys = Nothing
End Function

' ==================================================================
' = 概要    ファイル/フォルダパス一覧を取得する(Collection,Dirコマンド版)
' = 引数    sTrgtDir        String              [in]    対象フォルダ
' = 引数    cFileList       Object(Collection)  [out]   ファイル/フォルダパス一覧
' = 引数    lFileListType   Long                [in]    取得する一覧の形式
' =                                                         0：両方
' =                                                         1:ファイル
' =                                                         2:フォルダ
' =                                                         それ以外：格納しない
' = 引数    sFileExtStr     String              [in]    取得するファイルの拡張子(省略可能)
' =                                                       ex1) ""
' =                                                       ex2) "*"
' =                                                       ex3) "*.c"
' =                                                       ex4) "*.txt *.log *.csv"
' = 戻値    なし
' = 覚書    ・Dir コマンドによるファイル一覧取得。GetFileList() よりも高速。
' = 覚書    ・sFileExtStrはファイル指定時のみ有効
' = 依存    なし
' = 所属    Mng_FileSys.bas
' ==================================================================
Public Function GetFileListCmdClct( _
    ByVal sTrgtDir As String, _
    ByRef cFileList As Object, _
    ByVal lFileListType As Long, _
    Optional ByVal sFileExtStr As String = "" _
)
    Dim objFSO As Object 'FileSystemObjectの格納先
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Dir コマンド実行（出力結果を一時ファイルに格納）
    Dim sTmpFilePath As String
    Dim sExecCmd As String
    sTmpFilePath = CreateObject("WScript.Shell").CurrentDirectory & "\Dir.tmp"
    Dim sTrgtDirStr As String
    If sFileExtStr = "" Then
        sTrgtDirStr = """" & sTrgtDir & """"
    Else
        Dim vFileExtentions As Variant
        vFileExtentions = Split(sFileExtStr, " ")
        Dim lSplitIdx As Long
        For lSplitIdx = 0 To UBound(vFileExtentions)
            If sTrgtDirStr = "" Then
                sTrgtDirStr = """" & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            Else
                sTrgtDirStr = sTrgtDirStr & " """ & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            End If
        Next lSplitIdx
    End If
    Select Case lFileListType
        Case 0:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a > """ & sTmpFilePath & """"
        Case 1:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:a-d > """ & sTmpFilePath & """"
        Case 2:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:d > """ & sTmpFilePath & """"
        Case Else: sExecCmd = ""
    End Select
    With CreateObject("Wscript.Shell")
        .Run "cmd /c" & sExecCmd, 7, True
    End With
    
    Dim objFile As Object
    On Error Resume Next
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile(sTmpFilePath, 1)
        If Err.Number = 0 Then
            Do Until objFile.AtEndOfStream
                cFileList.Add objFile.ReadLine
            Loop
        Else
            MsgBox "ファイルが開けません: " & Err.Description
        End If
        Set objFile = Nothing   'オブジェクトの破棄
    Else
        MsgBox "エラー " & Err.Description
    End If
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    'オブジェクトの破棄
    On Error GoTo 0
End Function
    Private Sub Test_GetFileListCmdClct()
        Dim sRootDir As String
        sRootDir = "C:\codes"
        
        Dim cFileList As Object
        Set cFileList = CreateObject("System.Collections.ArrayList")
        
'        Call GetFileListCmdClct(sRootDir, cFileList, 0)
        Call GetFileListCmdClct(sRootDir, cFileList, 1)
'        Call GetFileListCmdClct(sRootDir, cFileList, 1, "*.c *.h")
'        Call GetFileListCmdClct(sRootDir, cFileList, 1, "*.vbs")
'        Call GetFileListCmdClct(sRootDir, cFileList, 1, "*")
'        Call GetFileListCmdClct(sRootDir, cFileList, 1, "")
'        Call GetFileListCmdClct(sRootDir, cFileList, 2)
        Stop
    End Sub

' ==================================================================
' = 概要    全てのマクロ/プロシージャをエクスポートする
' = 引数    bTargetBook     Workbook    [in]    エクスポート対象ブック
' = 戻値    なし
' = 覚書    ・以下の参照設定を追加する必要あり。
' =           - [ツール] -> [参照設定] ->「Microsoft Visual Basic for Applications Extensibility」
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Private Function ExportAllModules( _
    ByRef bTargetBook As Workbook _
)
    ' フォルダ作成
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim sExportDirPath As String
    sExportDirPath = bTargetBook.Path & "\" & bTargetBook.Name & ".bas"
    If Not objFSO.FolderExists(sExportDirPath) Then
        objFSO.CreateFolder (sExportDirPath)
    End If
    
    Debug.Print "*** Export all macros ***"
    Debug.Print "Target book : " & bTargetBook.Name
    Debug.Print "Export path : " & bTargetBook.Path
    Dim objModule As VBComponent
    For Each objModule In bTargetBook.VBProject.VBComponents
        ' モジュール種別判定
        Dim sExtension
        Select Case objModule.Type
            Case vbext_ct_ClassModule:  sExtension = "cls"
            Case vbext_ct_MSForm:       sExtension = "frm"
            Case vbext_ct_StdModule:    sExtension = "bas"
            Case vbext_ct_Document:     sExtension = "cls"
            Case Else:                  sExtension = ""
        End Select
        
        ' エクスポート実施
        Dim sExportDstFilePath
        sExportDstFilePath = sExportDirPath & "\" & objModule.Name & "." & sExtension
        If sExtension = "" Then
            Debug.Print "[Ignore  ] " & objModule.Name
        Else
            Call objModule.Export(sExportDstFilePath)
            Debug.Print "[Exported] " & objModule.Name & "." & sExtension
        End If
    Next
    Debug.Print ""
End Function

' ==================================================================
' = 概要    指定範囲のセルアドレス(e.g. A1)を結合した文字列を
' =         コピー(クリップボードに格納)する。
' =         例えば、B2～D2の範囲が指定された場合、"B2&C2&D2"をコピーする。
' = 引数    rRange          Range   [in]    セル範囲
' = 引数    bAbsRefRow      Boolean [in]    行に対する絶対参照指定 (省略可)
' = 引数    bAbsRefClm      Boolean [in]    列に対する絶対参照指定 (省略可)
' = 引数    bRefStyleR1C1   Boolean [in]    R1C1形式指定 (省略可)
' = 引数    sDelimiter      String  [in]    区切り文字 (省略可)
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Macros.bas
' ==================================================================
Private Function CopyConcatedCellAddresses( _
    ByRef rRange As Range, _
    Optional ByVal bAbsRefRow As Boolean = False, _
    Optional ByVal bAbsRefClm As Boolean = False, _
    Optional ByVal bRefStyleR1C1 = False, _
    Optional ByVal sDelimiter = "" _
)
    ' 範囲チェック
    If rRange.Columns.Count > 1 And rRange.Rows.Count > 1 Then
        MsgBox "[ERROR] 1行または1列を指定してください"
        Return
    End If
    
    ' セルアドレス取得＆結合
    Dim sConcatCellAdr As String
    sConcatCellAdr = ""
    Dim rCell As Range
    For Each rCell In rRange
        Dim sCellAdr As String
        Dim xlRefStyle As XlReferenceStyle
        If bRefStyleR1C1 = True Then
            xlRefStyle = xlR1C1
        Else
            xlRefStyle = xlA1
        End If
        sCellAdr = rCell.Address( _
            RowAbsolute:=bAbsRefRow, _
            ColumnAbsolute:=bAbsRefClm, _
            ReferenceStyle:=xlRefStyle _
        )
        Dim sDlmStr As String
        If sDelimiter = "" Then
            sDlmStr = "&"
        Else
            sDlmStr = "&""" & sDelimiter & """&"
        End If
        If sConcatCellAdr = "" Then
            sConcatCellAdr = sCellAdr
        Else
            sConcatCellAdr = sConcatCellAdr & sDlmStr & sCellAdr
        End If
    Next
    
    ' クリップボード設定
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText sConcatCellAdr
        .PutInClipboard
    End With
    MsgBox sConcatCellAdr
End Function


