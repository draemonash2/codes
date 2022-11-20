echo githubからダウンロードして比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\UpdateCheck.vbs" "%MYDIRPATH_CODES%" "https://github.com/draemonash2/codes/archive/master.zip" "codes-master"
echo %MYDIRPATH_CODES%内の_localフォルダを比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\DiffLocalDirs.vbs" "%MYDIRPATH_CODES%"

