@echo off
if {%MYPATH_CODES%} == {0} (
	echo target environment variable is nothing!
	pause
	exit /B 0
)
set sep_num=15
set inp_path=%MYPATH_CODES%\ruby\inp\input_study_ap.csv
ruby %MYPATH_CODES%\ruby\cre_studyPlan.rb %sep_num% %inp_path%
pause
