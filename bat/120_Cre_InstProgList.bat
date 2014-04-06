@echo off

call lib\010_Def_Datetime.bat

set OUT_PATH=Z:\100_Documents\100_PC\90_BackUp\InstallProgramList.txt

echo ###########################################   >  %OUT_PATH%
echo ### Execute Date is %datetime%                >> %OUT_PATH%
echo ###########################################   >> %OUT_PATH%
echo.                                              >> %OUT_PATH%

echo *******************************************   >> %OUT_PATH%
echo *** Destination : C:\prg                      >> %OUT_PATH%
echo *******************************************   >> %OUT_PATH%
cd "c:\prg"
dir                                                >> %OUT_PATH%
echo.                                              >> %OUT_PATH%

echo *******************************************   >> %OUT_PATH%
echo *** Destination : C:\Program Files            >> %OUT_PATH%
echo *******************************************   >> %OUT_PATH%
cd "C:\Program Files"
dir                                                >> %OUT_PATH%
echo.                                              >> %OUT_PATH%

echo *******************************************   >> %OUT_PATH%
echo *** Destination : C:\Program Files (x86)      >> %OUT_PATH%
echo *******************************************   >> %OUT_PATH%
cd "C:\Program Files (x86)"
dir                                                >> %OUT_PATH%
echo.                                              >> %OUT_PATH%
