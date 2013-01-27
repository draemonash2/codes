@echo off

echo ########### Create Tag Files #############
set /p DIR_PATH="###  select target path      : "
set /p TAG_TYPE="###  select tag type [a/c/g] : "
echo ###  a:all c:ctags g:gtags
echo ###  Create tag sourse is ...
echo ###    %DIR_PATH%
pause
echo ###  Wait for a while...
cd %DIR_PATH%
   if     %TAG_TYPE% == a (
    ctags -R
    gtags -v
) else if %TAG_TYPE% == c (
    ctags -R
) else if %TAG_TYPE% == g (
    gtags -v
) else (
    echo "###  error! select collect path type!
)
echo ############### Finish! ##################
pause
