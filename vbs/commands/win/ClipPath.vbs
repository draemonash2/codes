option explicit

'指定されたファイルのファイルパスをコピーする

if wscript.arguments.count > 0 then
	dim sPath
	dim sOut
	for each sPath in wscript.arguments
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
