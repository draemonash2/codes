Option Explicit

' 本スクリプトを自動実行させるためには、事前に以下の操作を行っておくこと！
' 
' 1. イベントビューアーを開く。
' 	コントロールパネル => 管理ツール => イベントビューアー
' 2. 画面左側「ログ」ペインにて、以下を選択。
' 	アプリケーションとサービスログ => Microsoft => Windows => DriverFrameworks-UserMode => Operational
' 3. 画面右側「操作」ペインにて「ログを有効化」をクリック。

'★特記事項★
'   現状、vbs からマクロを実行させると、なぜか以下の現象が発生する。
'   ・マクロが二回実行される
'   ・マクロ実行がめちゃくちゃ遅い
'   そのため、現状は Excel ファイルを開くのみに留めておく。

Const MACRO_BOOK_PATH = "C:\Users\draem_000\Documents\Dropbox\100_Documents\143_【生活】＜衣食住＞食事／身体\320_【身体】体重管理.xlsm"
Const MACRO_NAME = "生データ入力()"

Dim objExcel
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Open MACRO_BOOK_PATH
