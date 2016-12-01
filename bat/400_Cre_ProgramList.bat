@echo off

set LOG_FILE_PATH="C:\Users\draem_000\Documents\GoogleDrive\100_Programs\program_list.txt"
set TRGT_PATH_01=C:\prg
set TRGT_PATH_02=C:\prg_exe
set TRGT_PATH_03=C:\Program Files
set TRGT_PATH_04=C:\Program Files (x86)

echo.> "%LOG_FILE_PATH%"
echo ######################################################################### >> "%LOG_FILE_PATH%"
echo ### program list @ %date% %time% >> "%LOG_FILE_PATH%"
echo ######################################################################### >> "%LOG_FILE_PATH%"

echo create program folder list...
echo ============================================= >> "%LOG_FILE_PATH%"
echo = program folder list >> "%LOG_FILE_PATH%"
echo ============================================= >> "%LOG_FILE_PATH%"
	echo ### target folder "%TRGT_PATH_01%" ### >> "%LOG_FILE_PATH%" && dir /b /a:d "%TRGT_PATH_01%" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
	echo ### target folder "%TRGT_PATH_02%" ### >> "%LOG_FILE_PATH%" && dir /b /a:d "%TRGT_PATH_02%" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
	echo ### target folder "%TRGT_PATH_03%" ### >> "%LOG_FILE_PATH%" && dir /b /a:d "%TRGT_PATH_03%" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
	echo ### target folder "%TRGT_PATH_04%" ### >> "%LOG_FILE_PATH%" && dir /b /a:d "%TRGT_PATH_04%" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"

echo create program exe list...
echo ============================================= >> "%LOG_FILE_PATH%"
echo = program exe list >> "%LOG_FILE_PATH%"
echo ============================================= >> "%LOG_FILE_PATH%"
	echo ### target folder "%TRGT_PATH_01%" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "%TRGT_PATH_01%\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
	echo ### target folder "%TRGT_PATH_02%" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "%TRGT_PATH_02%\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
	echo ### target folder "%TRGT_PATH_03%" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "%TRGT_PATH_03%\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
	echo ### target folder "%TRGT_PATH_04%" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "%TRGT_PATH_04%\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"

::	cmd.exe /c "%LOG_FILE_PATH%"
