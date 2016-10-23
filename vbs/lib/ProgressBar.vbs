Option Explicit

Class ProgressBar
	Dim gobjExplorer
	Dim gsProgMsg
	Dim glProg100
	Dim glProg10
	Dim glStartTime
	
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
		
		gsProgMsg = ""
		glProg100 = 0
		glProg10 = 0
		glStartTime = 0
		
		Set gobjExplorer = CreateObject("InternetExplorer.Application")
		gobjExplorer.Navigate "about:blank"
		gobjExplorer.ToolBar = 0
		gobjExplorer.StatusBar = 0
		gobjExplorer.Width = 450
		gobjExplorer.Height = ( 28 * 2 ) + 65
		gobjExplorer.Left = ( intHorizontal - gobjExplorer.Width ) / 2
		gobjExplorer.Top = ( intVertical - gobjExplorer.Height ) / 2
		gobjExplorer.Visible = 1
		
		Call ActiveIE
		gobjExplorer.Document.Body.Style.Cursor = "wait"
		gobjExplorer.Document.Title = "進捗状況"
		
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
		Dim lBrNum
		Dim lLineNum
		'ウィンドウの高さ算出
		lBrNum = ( Len( sProgMsg ) - Len( Replace( sProgMsg, vbNewLine, "" ) ) ) / 2
		lLineNum = ( lBrNum + 1 ) + 4
		gobjExplorer.Height = ( 28 * lLineNum ) + 65
		
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
			sProgMsg = Replace( gsProgMsg, vbNewLine, "<br>" ) & "<br><br>"
		End If
		
		If glStartTime = 0 Then
			glStartTime = Now()
		Else
			'Do Nothing
		End If
		
		gobjExplorer.Document.Body.InnerHTML = _
			"<font face=""ＭＳ ゴシック"">" & _
			"<span style=""font-size:18px; line-height:22px;"">" & _
			sProgMsg & "処理中...<br>" & _
			DateDiff( "s", glStartTime, Now() ) & " [s] 経過...<br>" & _
			String( glProg10, "■") & String( 10 - glProg10, "□") & "  " & glProg100 & "% 完了" & _
			"</span>" & _
			"</font>" & _
			""
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
'Private Sub TestCase
'	Dim oProgBar
'	Dim lTestCase
'	Dim i
'	Dim iBefore
'	Dim iAfter
'	
'	lTestCase = 2
'	
'	Set oProgBar = New ProgressBar
'	With oProgBar
'		Select Case lTestCase
'			Case 1
'				.SetMsg( "Test Maggage" )
'				iBefore = Timer()
'				For i = 0 to 100
'					.SetProg( i )
'					WScript.Sleep 10
'				Next
'				iAfter = Timer()
'				MsgBox iAfter - iBefore
'			Case 2
'				.SetMsg( "Test Maggage" )
'				iBefore = Timer()
'				For i = 400 to 500
'					.SetProg( .ConvProgRange( 400, 500, i ) )
'					WScript.Sleep 10
'				Next
'				iAfter = Timer()
'				MsgBox iAfter - iBefore
'			Case Else
'				
'		End Select
'		.Quit()
'	End With
'End Sub

