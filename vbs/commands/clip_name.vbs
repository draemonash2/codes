option explicit

'�w�肳�ꂽ�t�@�C���̃t�@�C�������R�s�[����

if wscript.arguments.count > 0 then
	dim sPath
	dim sOut
	dim asSplitWords
	for each sPath in wscript.arguments
		asSplitWords = Split(sPath, "\")
		sPath = asSplitWords(UBound(asSplitWords))
		If sOut = "" then
			sOut = sPath
		else
			sOut = sOut & vbNewLine & sPath
		end if
	next
	
	With CreateObject("Wscript.Shell")
		With .Exec("clip")
			Call .StdIn.Write( sOut )
		End With
	End With
end if
