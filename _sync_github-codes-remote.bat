for /f "tokens=1,2,3" %%a in (%MYDIRPATH_CODES_CONFIG%\_sync_github-codes-remote.config) do (
	set HOST=%%a
	set USER=%%b
	set PASSWORD=%%c
)

echo githubからダウンロードして比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\SyncGithubToCodes.vbs" "%~dp0" "https://github.com/draemonash2/codes/archive/master.zip" "codes-master"

echo %MYDIRPATH_CODES%内の_localフォルダを比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\SyncCodesToLocal.vbs" "%~dp0"

echo %MYDIRPATH_CODES%とremote接続先のファイルを比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\SyncCodesToRemote.vbs" "%USER%" "%PASSWORD%" "%HOST%" "/home/%USER%"
