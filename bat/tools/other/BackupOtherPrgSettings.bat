@echo off

setlocal ENABLEDELAYEDEXPANSION

:: �f�B���N�g��
set OPT=
set OPT=!OPT! /MIR
set OPT=!OPT! /R:5
set OPT=!OPT! /W:30
set OPT=!OPT! /SL
set OPT=!OPT! /XD "System Volume Information"
robocopy "%USERPROFILE%\AppData\Roaming\Team Hasebe\TVClock"	"G:\�}�C�h���C�u\100_programs\120_setting\TVClock\Team Hasebe\TVClock" %OPT%
robocopy "%USERPROFILE%\AppData\Local\TresGrep"	"G:\�}�C�h���C�u\100_programs\120_setting\TresGrep\TresGrep" %OPT%

:: �t�@�C��
set OPT=
set OPT=!OPT! /R:5
set OPT=!OPT! /W:30
set OPT=!OPT! /SL
set OPT=!OPT! /XD "System Volume Information"
robocopy "%USERPROFILE%\AppData\Roaming\Microsoft\Templates"	"C:\codes\vba\word\AddIns" Normal.dotm %OPT%
robocopy "%USERPROFILE%\AppData\Local\TresGrep\TresGrep.exe_Url_o1daaqk3h25533o51axidnnzwhjzviq5\1.20.2019.1214"	"G:\�}�C�h���C�u\100_programs\120_setting\TresGrep" user.config %OPT%

endlocal
