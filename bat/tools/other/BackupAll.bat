@echo off

set CUR_DIR=%~dp0
	call "%CUR_DIR%\BackupPrograms.bat"
	call "%CUR_DIR%\BackupPrgExe.bat"
	call "%CUR_DIR%\BackupCodesSample.bat"
	call "%CUR_DIR%\BackupLibrary.bat"
	call "%CUR_DIR%\BackupRoot.bat"
	call "%CUR_DIR%\BackupSSHKey.bat"

pause

