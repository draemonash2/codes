Option Explicit

Function OutputAllElement2Console( _
	ByRef asOutTrgtArray _
)
	Dim lIdx
	Dim sOutStr
	sOutStr = "EleNum :" & Ubound( asOutTrgtArray ) + 1
	For lIdx = 0 to UBound( asOutTrgtArray )
		sOutStr = sOutStr & vbNewLine & asOutTrgtArray(lIdx)
	Next
	WScript.Echo sOutStr
End Function

Function OutputAllElement2LogFile( _
	ByRef asOutTrgtArray _
)
	Dim lIdx
	Dim objLogFile
	Dim sLogFilePath
	Dim objWshShell
	
	Set objWshShell = WScript.CreateObject( "WScript.Shell" )
	sLogFilePath = objWshShell.CurrentDirectory & "\" & replace( WScript.ScriptName, ".vbs", ".log" )
	Set objLogFile = CreateObject("Scripting.FileSystemObject").OpenTextFile( sLogFilePath, 2, True )
	
	objLogFile.WriteLine "EleNum :" & Ubound( asOutTrgtArray ) + 1
	For lIdx = 0 to UBound( asOutTrgtArray )
		objLogFile.WriteLine asOutTrgtArray( lIdx )
	Next
	objLogFile.Close
	
	Set objWshShell = Nothing
	Set objLogFile = Nothing
End Function

'Call Test
Private Sub Test
	Dim asFileList()
	Redim asFileList(3)

	asFileList(0) = 1
	asFileList(1) = 0
	asFileList(2) = 1
	asFileList(3) = 0
'	Call OutputAllElement2LogFile(asFileList)
	Call OutputAllElement2Console(asFileList)
End Sub
