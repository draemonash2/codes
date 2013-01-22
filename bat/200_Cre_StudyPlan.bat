@echo off
set sep_num=15
set inp_path=..\ruby\inp\input_study_ap.csv
ruby ..\ruby\cre_studyPlan.rb %sep_num% %inp_path%
pause
