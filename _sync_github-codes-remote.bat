@echo off
for /f "tokens=1,2,3,4" %%a in (%MYDIRPATH_CODES_CONFIG%\_sync_github-codes-remote.config) do (
	set HOST=%%a
	set USER=%%b
	set PASSWORD=%%c
	set HOMEDIR=%%d
)

echo github����_�E�����[�h���Ĕ�r���܂��B
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\UpdateCheck.vbs" "%MYDIRPATH_CODES%" "https://github.com/draemonash2/codes/archive/master.zip" "codes-master"
echo %MYDIRPATH_CODES%����_local�t�H���_���r���܂��B
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\DiffLocalDirs.vbs" "%MYDIRPATH_CODES%"
echo %MYDIRPATH_CODES%��remote�ڑ���̃t�@�C�����r���܂��B
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\DiffLclVsRmt.vbs" "%USER%" "%PASSWORD%" "%HOST%" "%HOMEDIR%"
