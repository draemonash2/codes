@echo off
setlocal ENABLEDELAYEDEXPANSION

	set LOG_FILE_PATH=%USERPROFILE%\Documents\Amazon Drive\100_Programs\program_list.txt
	echo.> "%LOG_FILE_PATH%"

	echo %date% %time%>> "%LOG_FILE_PATH%"
	echo #########################################################################>> "%LOG_FILE_PATH%"
	echo ### program list>> "%LOG_FILE_PATH%"
	echo #########################################################################>> "%LOG_FILE_PATH%"
	
	echo create program list (folder)...
	echo =============================================>> "%LOG_FILE_PATH%"
	echo = program list (folder)>> "%LOG_FILE_PATH%"
	echo =============================================>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\prg&&					echo ## target folder "!TRGT_PATH!">> "%LOG_FILE_PATH%" && dir /b /a:d "!TRGT_PATH!" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\prg_exe&&				echo ## target folder "!TRGT_PATH!">> "%LOG_FILE_PATH%" && dir /b /a:d "!TRGT_PATH!" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\Program Files&&		echo ## target folder "!TRGT_PATH!">> "%LOG_FILE_PATH%" && dir /b /a:d "!TRGT_PATH!" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\Program Files (x86)&&	echo ## target folder "!TRGT_PATH!">> "%LOG_FILE_PATH%" && dir /b /a:d "!TRGT_PATH!" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
	
	echo create program list (exe)...
	echo =============================================>> "%LOG_FILE_PATH%"
	echo = program list (exe)>> "%LOG_FILE_PATH%"
	echo =============================================>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\prg&&					echo ## target folder "!TRGT_PATH!">> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\prg_exe&&				echo ## target folder "!TRGT_PATH!">> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\Program Files&&		echo ## target folder "!TRGT_PATH!">> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\Program Files (x86)&&	echo ## target folder "!TRGT_PATH!">> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"

::	cmd.exe /c "%LOG_FILE_PATH%"
endlocal
