Option Explicit

Class LogMng
	Dim gbLogFileEnable
	Dim goLogFile
	
	Private Sub Class_Initialize()
		Call LogInit()
	End Sub
	Private Sub Class_Terminate()
		Call LogFileClose()
	End Sub
	
	Private Function LogInit()
		gbLogFileEnable = False
		Set goLogFile = Nothing
	End Function
	
	 '第二引数：IOモード（"r":読出し、"w":新規書込み、"+w":追加書込み）
	Public Function LogFileOpen( _
		ByVal sTrgtPath, _
		ByVal sIOMode _
	)
		Dim lIOMode
		Select Case LCase( sIOMode )
			Case "r" :	lIOMode = 1
			Case "w" :	lIOMode = 2
			Case "+w" :	lIOMode = 8
			Case Else :	Exit Function
		End Select
		
		Set goLogFile = CreateObject("Scripting.FileSystemObject").OpenTextFile( sTrgtPath, lIOMode, True)
		gbLogFileEnable = True
	End Function
	
	Public Function LogPuts( _
		ByVal sMsg _
	)
		If gbLogFileEnable = True Then
			goLogFile.WriteLine sMsg
		Else
			WScript.Echo sMsg
		End If
	End Function
	
	Public Function LogFileClose()
		If gbLogFileEnable = True Then
			goLogFile.Close
			gbLogFileEnable = False
		Else
			'Do Nothing
		End If
	End Function
End Class
'	Call Test
'	Private Sub Test
'		Dim oLog
'		Set oLog = New LogMng
'		
'		Call oLog.LogFileOpen( _
'			"C:\Users\draem_000\Desktop\test.log", _
'			"+w" _
'		)
'		oLog.LogPuts "desu"
'		oLog.LogPuts "yorosiku"
'		Call oLog.LogFileClose
'		
'		Call oLog.LogFileOpen( _
'			"C:\Users\draem_000\Desktop\test2.log", _
'			"w" _
'		)
'		oLog.LogPuts "desu"
'		oLog.LogPuts "yorosiku"
'		oLog.LogPuts "you"
'		Call oLog.LogFileClose
'		
'		oLog.LogPuts "desu"
'		oLog.LogPuts "yorosiku"
'		Call oLog.LogFileClose
'		
'		Set oLog = Nothing
'	End Sub
