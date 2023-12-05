@echo off
if "%1" == "" (
	echo "[error] wrong arguments"
	echo "    [usage] exec_ps_script.bat ps1_script_path.ps1"
) else (
	powershell -NoProfile -ExecutionPolicy Unrestricted "%1" "%2" "%3" "%4" "%5" "%6" "%7" "%8" "%9"
)
