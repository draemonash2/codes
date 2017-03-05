Option Explicit

Public Function UpdateIniFile( _
	ByVal sTrgtIniPath, _
	ByVal sItemName, _
	ByVal sItemPath, _
	ByVal sItemType, _
	ByVal sItemIcon, _
	ByVal sItemExt _
)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim objFileOpenAsWrite
	Dim lLineIdx
	
	If objFSO.FileExists( sTrgtIniPath ) Then
		Dim vTextAll
		vTextAll = TextFile2VariantArray( sTrgtIniPath )
		
		Dim lTailIdx
		lTailIdx = GetTailIdx( vTextAll )
		
		Set objFileOpenAsWrite = objFSO.OpenTextFile( sTrgtIniPath, 2, True )
		
		For lLineIdx = 0 To Ubound( vTextAll )
			If lLineIdx = 1 Then
				objFileOpenAsWrite.WriteLine "Count=" & lTailIdx + 1
			Else
				objFileOpenAsWrite.WriteLine vTextAll( lLineIdx )
			End If
		Next
		
		objFileOpenAsWrite.WriteLine "Name" & lTailIdx & "=" & sItemName
		objFileOpenAsWrite.WriteLine "Path" & lTailIdx & "=" & sItemPath
		objFileOpenAsWrite.WriteLine "Type" & lTailIdx & "=" & sItemType
		objFileOpenAsWrite.WriteLine "Icon" & lTailIdx & "=" & sItemIcon
		objFileOpenAsWrite.WriteLine "Ext" & lTailIdx & "=" & sItemExt
		
		objFileOpenAsWrite.Close
		Set objFileOpenAsWrite = Nothing
	Else
		Set objFileOpenAsWrite = objFSO.OpenTextFile( sTrgtIniPath, 2, True )
		
		objFileOpenAsWrite.WriteLine "[X-Finder]"
		objFileOpenAsWrite.WriteLine "Count=1"
		objFileOpenAsWrite.WriteLine "Name0=" & sItemName
		objFileOpenAsWrite.WriteLine "Path0=" & sItemPath
		objFileOpenAsWrite.WriteLine "Type0=" & sItemType
		objFileOpenAsWrite.WriteLine "Icon0=" & sItemIcon
		objFileOpenAsWrite.WriteLine "Ext0=" & sItemExt
		
		objFileOpenAsWrite.Close
		Set objFileOpenAsWrite = Nothing
	End If
End Function
'	Call Test_UpdateIniFile()
	Private Sub Test_UpdateIniFile()
		Dim sTrgtPath
		sTrgtPath = "C:\Users\draem_000\Desktop\data\vbs\lib\test.ini"
		
		Call UpdateIniFile( _
			sTrgtPath, _
			"aaaa", _
			"c:\test", _
			5, _
			"", _
			"" _
		)
		Call UpdateIniFile( _
			sTrgtPath, _
			"bbb", _
			"c:\tests", _
			5, _
			"", _
			"" _
		)
	End Sub

Private Function GetTailIdx( _
	ByRef vTextAll _
)
	Dim sTailLine
	sTailLine = vTextAll( Ubound( vTextAll ) )
	If InStr( sTailLine, "Ext" ) > 0 Then
		GetTailIdx = CLng( Replace( Replace( sTailLine, "Ext", "" ), "=", "" ) ) + 1
	Else
		GetTailIdx = 0
	End If
End Function
'	Call Test_GetTailIdx()
	Private Sub Test_GetTailIdx()
		Dim sTrgtPath
		sTrgtPath = "C:\Users\draem_000\Desktop\data\vbs\lib\_fav_program_ÅyVideoÅzRecord.ini"
		
		Dim vTextAll
		vTextAll = TextFile2VariantArray( sTrgtPath )
		MsgBox vTextAll( Ubound( vTextAll ) )
		
		Dim sOutStr
		sOutStr = ""
		sOutStr = sOutStr & vbNewLine & "*** test start! ***"
		sOutStr = sOutStr & vbNewLine & GetTailIdx( vTextAll )
		sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
		MsgBox sOutStr
	End Sub

Private Function TextFile2VariantArray( _
	ByVal sTrgtPath _
)
	Dim objFile
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile( sTrgtPath, 1 )
	Dim sTextAll
	sTextAll = objFile.ReadAll
	Dim vTextAll
	vTextAll = Split( sTextAll, vbNewLine )
	If Ubound( vTextAll ) > 1 Then
		ReDim Preserve vTextAll( UBound( vTextAll ) - 1 )
	Else
		'Do Nothing
	End If
	TextFile2VariantArray = vTextAll
	objFile.Close
	Set objFile = Nothing
End Function
'	Call Test_TextFile2VariantArray()
	Private Sub Test_TextFile2VariantArray()
		Dim sTrgtPath
		sTrgtPath = "C:\Users\draem_000\Desktop\data\vbs\lib\_fav_program_ÅyVideoÅzRecord.ini"
		
		Dim sOutStr
		sOutStr = ""
		sOutStr = sOutStr & vbNewLine & "*** test start! ***"
		Dim vTextAll
		vTextAll = TextFile2VariantArray( sTrgtPath )
		sOutStr = sOutStr & vbNewLine & Ubound( vTextAll )
		Dim lIdx
		For lIdx = 0 to UBound( vTextAll )
			sOutStr = sOutStr & vbNewLine & vTextAll(lIdx)
		Next
		sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
		MsgBox sOutStr
	End Sub
