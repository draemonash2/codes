@echo off

echo ########### shutdown later #############
set /p SELECTED_UNIT="###  please select unit [h/m/s] : "
set /p TIME_VALUE="###  please set time value : "
if %SELECTED_UNIT% == h (
    set /p ANSWER="###  shutdown in %TIME_VALUE% hour. ok? [y/n] : "
    pause
) else if %SELECTED_UNIT% == m (
    set /p ANSWER="###  shutdown in %TIME_VALUE% minite. ok? [y/n] : "
    pause
) else if %SELECTED_UNIT% == s (
    set /p ANSWER="###  shutdown in %TIME_VALUE% second. ok? [y/n] : "
    pause
) else (
    goto ERROR
)

if %ANSWER% == y (
    rem
) else (
    goto ERROR
)

if %SELECTED_UNIT% == h (
    set /a "SET_VALUE = TIME_VALUE * 3600"
) else if %SELECTED_UNIT% == m (
    set /a "SET_VALUE = TIME_VALUE * 60"
) else if %SELECTED_UNIT% == s (
    set /a "SET_VALUE = TIME_VALUE"
) else (
    goto ERROR
)

shutdown /s /t %SET_VALUE%
goto END

:ERROR
echo ###  processing was interrupted.
goto END

:END
echo ############### finish! ##################
pause
