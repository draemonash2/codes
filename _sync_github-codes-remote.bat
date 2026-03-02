for /f "tokens=1,2,3" %%a in (%MYDIRPATH_CODES_CONFIG%\_ssh_target_a.config) do (
	set HOST=%%a
	set USER=%%b
	set PASSWORD=%%c
)

echo githubからダウンロードして比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\DiffCodes_Lcl-vs-Github.vbs" "%~dp0" "https://github.com/draemonash2/codes/archive/master.zip" "codes-master"

echo %MYDIRPATH_CODES%内の_localフォルダを比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\DiffDirs_Xxx-vs-XxxLcl.vbs" "%~dp0"

echo %MYDIRPATH_CODES%とremote接続先のファイルを比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\DiffDotfiles_LclWin-vs-RmtLinux.vbs" "%USER%" "%PASSWORD%" "%HOST%"
