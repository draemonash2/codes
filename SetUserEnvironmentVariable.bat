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
	setx MYPATH_CODE_BAT	"C:\codes\bat"
	setx MYPATH_CODE_C		"C:\codes\c"
	setx MYPATH_CODE_HTTP	"C:\codes\http"
	setx MYPATH_CODE_JAVA	"C:\codes\java"
	setx MYPATH_CODE_PYTHON	"C:\codes\python"
	setx MYPATH_CODE_RUBY	"C:\codes\ruby"
	setx MYPATH_CODE_SH		"C:\codes\sh"
	setx MYPATH_CODE_VBA	"C:\codes\vba"
	setx MYPATH_CODE_VBS	"C:\codes\vbs"
	setx MYPATH_CODE_VDM	"C:\codes\vdm++"
	echo In order to reflect this setting, please restart the windows!
) else if %ANS% == n (
	reg delete HKCU\Environment /v MYPATH_CODE_BAT
	reg delete HKCU\Environment /v MYPATH_CODE_C
	reg delete HKCU\Environment /v MYPATH_CODE_HTTP
	reg delete HKCU\Environment /v MYPATH_CODE_JAVA
	reg delete HKCU\Environment /v MYPATH_CODE_PYTHON
	reg delete HKCU\Environment /v MYPATH_CODE_RUBY
	reg delete HKCU\Environment /v MYPATH_CODE_SH
	reg delete HKCU\Environment /v MYPATH_CODE_VBA
	reg delete HKCU\Environment /v MYPATH_CODE_VBS
	reg delete HKCU\Environment /v MYPATH_CODE_VDM
	echo In order to reflect this setting, please restart the windows!
) else (
	echo [error] illegal answer!!!
)

pause
