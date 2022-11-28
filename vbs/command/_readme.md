
# commandフォルダ登録基準

- メッセージ出力時、MsgBoxではなくWScript.Echoを使用していること
	- cscript経由での実行時、標準出力させるため
- 動作のパラメータは、対話形式ではなくコマンドライン引数や入力ファイルから指定できること
	- コマンド実行ごとに動作が止まらず、連続実行できるようにするため

# コマンドラインから実行する方法

- 「cscript //nologo」をつけて実行する。
	- 例） `cscript //nologo OutputShortcutTrgtPath.vbs -v C:\Users\test.txt.lnk`

