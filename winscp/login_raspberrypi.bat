@echo off
for /f "tokens=1,2,3" %%a in (%MYDIRPATH_CODES_CONFIG%\_ssh_target_a.config) do (
	set HOST=%%a
	set USER=%%b
	set PASSWORD=%%c
)
start %MYEXEPATH_WINSCP% sftp://%USER%:%PASSWORD%@%HOST%:22
