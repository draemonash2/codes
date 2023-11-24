@echo off

set CUR_DIR=%~dp0
	call "%CUR_DIR%\BackupPrograms.bat"
	call "%CUR_DIR%\BackupCodesSample.bat"
	call "%CUR_DIR%\BackupLibrary.bat"

pause

