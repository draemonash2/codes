@echo off
if not {%MYPATH_CODE_RUBY%} == {0} (
	echo target environment variable is nothing!
	pause
	exit /B 0
)
set sep_num=15
set inp_path=%MYPATH_CODE_RUBY%\inp\input_study_ap.csv
ruby %MYPATH_CODE_RUBY%\cre_studyPlan.rb %sep_num% %inp_path%
pause
