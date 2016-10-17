Option Explicit

Function OutputAllElement( _
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
