Option Explicit

Class ProgressBar
	Dim gobjExplorer
	Dim gsProgMsg
	Dim glProg100
	Dim glProg10
	Dim glStartTime
	Dim goIE
	
	Private Sub Class_Initialize()
		Dim sIncludeVbsPath
		sIncludeVbsPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" ) & "\IE.vbs"
		ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile( sIncludeVbsPath ).ReadAll()
		
		Set goIE = New IE
		gsProgMsg = ""
		glProg100 = 0
		glProg10 = 0
		glStartTime = 0
		
		goIE.Title = "進捗状況"
		goIE.Activate
		
		SetProg(0)
	End Sub
	
	Private Sub Class_Terminate()
		'Do Nothing
	End Sub
	
	' メッセージを指定
	' ★注意★
	'   本関数は若干処理時間がかかります。
	'   一定時間感覚を空けて呼び出すこと。
	Public Function SetMsg( _
		ByVal sProgMsg _
	)
		gsProgMsg = sProgMsg
		
		PutProg()
	End Function
	
	' 0 ～ 100 を指定
	' ★注意★
	'   本関数は若干処理時間がかかります。
	'   一定時間感覚を空けて呼び出すこと。
	Public Function SetProg( _
		ByVal lProg100 _
	)
		If lProg100 > 100 Or lProg100 < 0 Then
			MsgBox "指定されたプログレスバーの進捗が最大最小範囲外の値が指定されています！" & vbNewLine & _
				   "値：" & lProg100
			MsgBox "プログラムを中止します！"
			Call Quit
			WScript.Quit
		End If
		
		glProg100 = Fix( lProg100 )
		glProg10 = Fix( lProg100 / 10 )
		
		PutProg()
	End Function
	
	' 進捗を変換（例：0～500 を 0～100 に変換）
	Public Function ConvProgRange( _
		ByVal lInMin, _
		ByVal lInMax, _
		ByVal lInProg _
	)
		Dim lConvMax
		Dim lConvProg
		
		If ( lInMin >= 0 And lInMax >= 0 And lInProg >= 0 ) And _
		   ( lInMax >= lInProg And lInProg >= lInMin And lInMax >= lInMin ) Then
			'Do Nothing
		Else
			MsgBox "ConvProgRange関数の引数が不正です。" & vbNewLine & _
				   "lInMin  : " & lInMin & vbNewLine & _
				   "lInMax  : " & lInMax & vbNewLine & _
				   "lInProg : " & lInProg
			MsgBox "プログラムを中断します。"
			WScript.Quit()
		End If
		
		lConvMax = ( lInMax - lInMin ) + 1
		lConvProg = ( lInProg - lInMin ) + 1
		ConvProgRange = ( lConvProg / lConvMax ) * 100
	End Function
	
	Private Function PutProg()
		Dim sProgMsg
		If gsProgMsg = "" Then
			sProgMsg = gsProgMsg
		Else
			sProgMsg = gsProgMsg & String( 2, vbNewLine )
		End If
		sProgMsg = _
			sProgMsg & "処理中..."& vbNewLine & _
			DateDiff( "s", glStartTime, Now() ) & " [s] 経過..." & vbNewLine & _
			String( glProg10, "■") & String( 10 - glProg10, "□") & "  " & glProg100 & "% 完了"
		goIE.Text = sProgMsg
		
		If glStartTime = 0 Then
			glStartTime = Now()
		Else
			'Do Nothing
		End If
	End Function
	
	Public Function Quit()
		goIE.Quit
	End Function
End Class
'	Private Sub Test
'		Dim oProgBar
'		Dim lTestCase
'		Dim i
'		Dim iBefore
'		Dim iAfter
'		
'		lTestCase = 1
'		
'		Set oProgBar = New ProgressBar
'		With oProgBar
'			Select Case lTestCase
'				Case 1
'					.SetMsg( "Test Message" )
'					iBefore = Timer()
'					For i = 0 to 100
'						.SetProg( i )
'						WScript.Sleep 10
'					Next
'					iAfter = Timer()
'					MsgBox iAfter - iBefore
'				Case 2
'					.SetMsg( "Test Message" )
'					iBefore = Timer()
'					For i = 400 to 500
'						.SetProg( .ConvProgRange( 400, 500, i ) )
'						WScript.Sleep 10
'					Next
'					iAfter = Timer()
'					MsgBox iAfter - iBefore
'				Case Else
'					
'			End Select
'			.Quit()
'		End With
'	End Sub
'	Call Test
