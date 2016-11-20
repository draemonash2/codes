:: *******************************************************************
:: * 目的：スクリプトがカレントディレクトリを取得する際、ローカル権限
:: *       における実行と管理者権限における実行とでは返却値が異なる
:: *       場合がある。その場合、事前にユーザー環境変数に設定しておく
:: *       ことで、権限毎の実行結果を合わせることができる。
:: *       
:: *       本バッチファイルでは、ユーザー環境変数の設定と削除を自動化
:: *       することを目的とする。
:: *******************************************************************
@echo off

echo add or delete user environment variable?
set /p ANS="  add=>y, delete=>n : "
if %ANS% == y (
	setx MYPATH_CODES	"C:\codes"
	echo In order to reflect this setting, please restart the windows!
) else if %ANS% == n (
	reg delete HKCU\Environment /v MYPATH_CODES
	echo In order to reflect this setting, please restart the windows!
) else (
	echo [error] illegal answer!!!
)

pause
