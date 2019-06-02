@echo off

set sep_num=15
set ruby_path=%~dp0..\ruby
set inp_path=%ruby_path%\inp\input_study_ap.csv
set script_path=%ruby_path%\cre_studyPlan.rb

ruby %script_path% %sep_num% %inp_path%
pause
