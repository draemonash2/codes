Option Explicit

Class LogMng
	Private gbLogFileEnable
	Private goLogFile
	
	Private Sub Class_Initialize()
		Call Init()
	End Sub
	Private Sub Class_Terminate()
		Call Close()
	End Sub
	
	Private Function Init()
		gbLogFileEnable = False
		Set goLogFile = Nothing
	End Function
	
	'第一引数：出力モード（0:標準出力、1:ログファイル出力）
	'戻り値：出力モード変更 実行結果（True:成功、False:失敗）
	Public Function Mode( _
		ByVal lSelectedMode _
	)
		If lSelectedMode = 0 Then
			gbLogFileEnable = False
			Mode = True
		ElseIf lSelectedMode = 1 Then
			gbLogFileEnable = True
			Mode = True
		Else
			Mode = False
		End If
	End Function
	
	 '第二引数：IOモード（"r":読出し、"w":新規書込み、"+w":追加書込み）
	Public Function Open( _
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
	
	Public Function Puts( _
		ByVal sOutputMsg _
	)
		If gbLogFileEnable = True Then
			goLogFile.WriteLine sOutputMsg
		Else
			WScript.Echo sOutputMsg
		End If
	End Function
	
	Public Function Close()
		If gbLogFileEnable = True Then
			goLogFile.Close
			gbLogFileEnable = False
		Else
			'Do Nothing
		End If
	End Function
End Class
	If WScript.ScriptName = "Log.vbs" Then
		Call Test_LogMng
	End If
	Private Sub Test_LogMng
		Dim oLog1
		Set oLog1 = New LogMng
		Dim oLog2
		Set oLog2 = New LogMng
		Call oLog1.Open( Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" ) & "\LogTest1.log", "+w" )
		Call oLog2.Open( Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" ) & "\LogTest2.log", "+w" )
		
		oLog1.Puts "desu"
		oLog1.Puts "yorosiku"
		oLog2.Puts "desu"
		oLog2.Puts "yorosiku"
		oLog2.Puts "you"
		Call oLog2.Close
		oLog2.Puts "desu"
		oLog2.Puts "yorosiku"
		
		Call oLog1.Close
		Call oLog2.Close
		Set oLog1 = Nothing
		Set oLog2 = Nothing
	End Sub
