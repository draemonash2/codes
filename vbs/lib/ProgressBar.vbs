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
		gobjExplorer.Document.Title = "�i����"
		
		SetProg(0)
	End Sub
	
	Private Sub Class_Terminate()
		'Do Nothing
	End Sub
	
	' ���b�Z�[�W���w��
	' �����Ӂ�
	'   �{�֐��͎኱�������Ԃ�������܂��B
	'   ��莞�Ԋ��o���󂯂ČĂяo�����ƁB
	Public Function SetMsg( _
		ByVal sProgMsg _
	)
		Dim lBrNum
		Dim lLineNum
		'�E�B���h�E�̍����Z�o
		lBrNum = ( Len( sProgMsg ) - Len( Replace( sProgMsg, vbNewLine, "" ) ) ) / 2
		lLineNum = ( lBrNum + 1 ) + 4
		gobjExplorer.Height = ( 28 * lLineNum ) + 65
		
		gsProgMsg = sProgMsg
		
		PutProg()
	End Function
	
	' 0 �` 100 ���w��
	' �����Ӂ�
	'   �{�֐��͎኱�������Ԃ�������܂��B
	'   ��莞�Ԋ��o���󂯂ČĂяo�����ƁB
	Public Function SetProg( _
		ByVal lProg100 _
	)
		If lProg100 > 100 Or lProg100 < 0 Then
			MsgBox "�w�肳�ꂽ�v���O���X�o�[�̐i�����ő�ŏ��͈͊O�̒l���w�肳��Ă��܂��I" & vbNewLine & _
				   "�l�F" & lProg100
			MsgBox "�v���O�����𒆎~���܂��I"
			Call Quit
			WScript.Quit
		End If
		
		glProg100 = Fix( lProg100 )
		glProg10 = Fix( lProg100 / 10 )
		
		PutProg()
	End Function
	
	' �i����ϊ��i��F0�`500 �� 0�`100 �ɕϊ��j
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
			MsgBox "ConvProgRange�֐��̈������s���ł��B" & vbNewLine & _
				   "lInMin  : " & lInMin & vbNewLine & _
				   "lInMax  : " & lInMax & vbNewLine & _
				   "lInProg : " & lInProg
			MsgBox "�v���O�����𒆒f���܂��B"
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
			"<font face=""�l�r �S�V�b�N"">" & _
			"<span style=""font-size:18px; line-height:22px;"">" & _
			sProgMsg & "������...<br>" & _
			DateDiff( "s", glStartTime, Now() ) & " [s] �o��...<br>" & _
			String( glProg10, "��") & String( 10 - glProg10, "��") & "  " & glProg100 & "% ����" & _
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
		
		Const strIEexe = "iexplore.exe" 'IE�̃v���Z�X��
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

