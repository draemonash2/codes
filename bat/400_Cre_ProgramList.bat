@echo off
setlocal ENABLEDELAYEDEXPANSION

	set LOG_FILE_PATH=%USERPROFILE%\Documents\Amazon Drive\100_Programs\program_list.txt
	echo.> "%LOG_FILE_PATH%"

	echo ######################################################################### >> "%LOG_FILE_PATH%"
	echo ### program list @ %date% %time% >> "%LOG_FILE_PATH%"
	echo ######################################################################### >> "%LOG_FILE_PATH%"
	
	echo create program folder list...
	echo ============================================= >> "%LOG_FILE_PATH%"
	echo = program folder list >> "%LOG_FILE_PATH%"
	echo ============================================= >> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\prg&&					echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /a:d "!TRGT_PATH!" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 
		set TRGT_PATH=C:\prg_exe&&				echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /a:d "!TRGT_PATH!" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 
		set TRGT_PATH=C:\Program Files&&		echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /a:d "!TRGT_PATH!" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 
		set TRGT_PATH=C:\Program Files (x86)&&	echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /a:d "!TRGT_PATH!" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 
	
	echo create program exe list...
	echo ============================================= >> "%LOG_FILE_PATH%"
	echo = program exe list >> "%LOG_FILE_PATH%"
	echo ============================================= >> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\prg&&					echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\prg_exe&&				echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\Program Files&&		echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\Program Files (x86)&&	echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.exe" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%"

	echo ######################################################################### >> "%LOG_FILE_PATH%"
	echo ### shortcut list >> "%LOG_FILE_PATH%"
	echo ######################################################################### >> "%LOG_FILE_PATH%"
	
	echo create send to program list...
	echo ============================================= >> "%LOG_FILE_PATH%"
	echo = send to program list >> "%LOG_FILE_PATH%"
	echo ============================================= >> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\ProgramData\AppData\Roaming\Microsoft\Windows\SendTo&&			echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.lnk" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 
		set TRGT_PATH=C:\Users\Administrator\AppData\Roaming\Microsoft\Windows\SendTo&&	echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.lnk" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 
		set TRGT_PATH=%USERPROFILE%\AppData\Roaming\Microsoft\Windows\SendTo&&			echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.lnk" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 
	
	echo create start menu program list...
	echo ============================================= >> "%LOG_FILE_PATH%"
	echo = start menu program list >> "%LOG_FILE_PATH%"
	echo ============================================= >> "%LOG_FILE_PATH%"
		set TRGT_PATH=C:\ProgramData\Microsoft\Windows\Start Menu&&						echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.lnk" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 
		set TRGT_PATH=%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu&&		echo ### target folder "!TRGT_PATH!" ### >> "%LOG_FILE_PATH%" && dir /b /s /a:a-d "!TRGT_PATH!\*.lnk" >> "%LOG_FILE_PATH%" && echo.>> "%LOG_FILE_PATH%" 

::	cmd.exe /c "%LOG_FILE_PATH%"
endlocal
