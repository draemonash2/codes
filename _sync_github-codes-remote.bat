for /f "tokens=1,2,3" %%a in (%MYDIRPATH_CODES_CONFIG%\_sync_github-codes-remote.config) do (
	set HOST=%%a
	set USER=%%b
	set PASSWORD=%%c
)

echo github����_�E�����[�h���Ĕ�r���܂��B
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\SyncGithubToCodes.vbs" "%~dp0" "https://github.com/draemonash2/codes/archive/master.zip" "codes-master"

echo %MYDIRPATH_CODES%����_local�t�H���_���r���܂��B
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\SyncCodesToLocal.vbs" "%~dp0"

echo %MYDIRPATH_CODES%��remote�ڑ���̃t�@�C�����r���܂��B
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\SyncCodesToRemote.vbs" "%USER%" "%PASSWORD%" "%HOST%" "/home/%USER%"
