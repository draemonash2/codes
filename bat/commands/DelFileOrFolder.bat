@echo off

setlocal

set JDG_FILE=%~a1
set FILE=%1

::�t�@�C�����݊m�F
if not exist %FILE% goto end

::�t�@�C�������݂����ꍇ�A�폜
if %JDG_FILE:~0,1% == d (
	rmdir /S /Q %FILE%
	echo Directry %FILE% is Deleted!
) else if %JDG_FILE:~0,1%==- (
	del /Q /F %FILE%
	echo File %FILE% is Deleted!
) else (
	echo %FILE% is not File/Folder Path!
)

:end
endlocal
