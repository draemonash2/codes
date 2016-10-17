Option Explicit

Class ProgressBar
	Dim gobjExplorer
	Dim gsProgMsg
	Dim glProg100
	
	Private Sub Class_Initialize()
		Dim objWMIService
		Dim colItems
		Dim strComputer
		Dim objItem
		Dim intHorizontal
		Dim intVertical
		strComputer = "."
		Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
		For Each objItem in colItems
			intHorizontal = objItem.ScreenWidth
			intVertical = objItem.ScreenHeight
		Next
		Set objWMIService = Nothing
		Set colItems = Nothing
		
		Set gobjExplorer = CreateObject("InternetExplorer.Application")
		gobjExplorer.Navigate "about:blank"
		gobjExplorer.ToolBar = 0
		gobjExplorer.StatusBar = 0
		gobjExplorer.Left = (intHorizontal - 400) / 2
		gobjExplorer.Top = (intVertical - 200) / 2
		gobjExplorer.Width = 400
		gobjExplorer.Height = 10
		gobjExplorer.Visible = 1
		
		Call ActiveIE
		gobjExplorer.Document.Body.Style.Cursor = "wait"
		gobjExplorer.Document.Title = "進捗状況"
		
		glProg100 = 0
		gsProgMsg = ""
		
		SetProg(0)
	End Sub
	
	Private Sub Class_Terminate()
		'Do Nothing
	End Sub
	
	Public Function SetMsg( _
		ByVal sMessage _
	)
		gsProgMsg = Replace( sMessage, vbNewLine, "<br>" ) '改行文字を<br>に置換
		PutProg()
	End Function
	
	' 0 ～ 100 を指定
	Public Function SetProg( _
		ByVal lProg100 _
	)
		If lProg100 > 100 Or lProg100 < 0 Then
			MsgBox "プログレスバーの進捗に規定値[0-100]外の値が指定されています！" & vbNewLine & _
				   "値：" & lProg100
			MsgBox "プログラムを中止します！"
			Call Quit
			WScript.Quit
		End If
		
		glProg100 = Fix(lProg100)
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
		Dim lProg10
		Dim lLineNum
		Dim lBrNum
		lProg10 = Fix( glProg100 / 10 )
		lBrNum = ( Len(gsProgMsg) - Len(Replace(gsProgMsg, "<br>", "")) ) / 4
		lLineNum = ( lBrNum + 1 ) + 3
		gobjExplorer.Height = 5 + ( 40 * lLineNum )
		gobjExplorer.Document.Body.InnerHTML = _
			gsProgMsg & "<br>" & _
			"<br>" & _
			"処理中..." & "<br>" & _
			String( lProg10, "■") & String( 10 - lProg10, "□") & "  " & glProg100 & "% 完了"
	End Function
	
	Public Function Quit()
		gobjExplorer.Document.Body.Style.Cursor = "default"
		gobjExplorer.Quit
	End Function
	
	Private Function ActiveIE()
		Dim objWshShell
		Dim intProcID
	
		Const strIEexe = "iexplore.exe" 'IEのプロセス名
		intProcID = GetProcID(strIEexe)
		
		Set objWshShell = CreateObject("Wscript.Shell")
		objWshShell.AppActivate intProcID
		
		Set objWshShell = Nothing
	End Function
	
	Private Function GetProcID(ProcessName)
		Dim Service
		Dim QfeSet
		Dim Qfe
		Dim intProcID
		
		Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
		Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='"& ProcessName &"'")
		
		intProcID = 0
		For Each Qfe in QfeSet
			intProcID = Qfe.ProcessId
			GetProcID = intProcID
			Exit For
		Next
	End Function
End Class

'Call TestCase
Private Sub TestCase
	Dim oProgBar
	Dim lTestCase
	Dim i
	
	lTestCase = 0
	
	Set oProgBar = New ProgressBar
	With oProgBar
		Select Case lTestCase
			Case 1
				.SetMsg( "Test Maggage" )
				For i = 0 to 100
					.SetProg( i )
					WScript.Sleep 50
				Next
			Case 2
				.SetMsg( "Test Maggage" )
				For i = 100 to 500
					.SetProg( .ConvProgRange( 100, 500, i ) )
					WScript.Sleep 10
				Next
			Case Else
				MsgBox .ConvProgRange( 100, 500, 99 )
		End Select
		.Quit()
	End With
End Sub

