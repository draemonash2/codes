Option Explicit

' �萔�͎������ē����o��������
Const LINE_HEIGHT_RATIO = 1.22
Const WIN_LINE_HEIGHT_RATIO = 1.27
Const HEADER_HEIGHT = 65

Class IE
	Dim gobjExplorer
	Dim glHorizontal
	Dim glVertical
	Dim gsFont
	Dim glFontSize
	Dim glLineHeight
	
	Private Sub Class_Initialize()
		'��ʃT�C�Y�擾
		Dim objWMIService
		Dim colItems
		Dim objItem
		Set objWMIService = GetObject("Winmgmts:\\.\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
		For Each objItem in colItems
			glHorizontal = objItem.ScreenWidth
			glVertical = objItem.ScreenHeight
		Next
		Set objWMIService = Nothing
		Set colItems = Nothing
		
		gsFont = "�l�r �S�V�b�N"
		glFontSize = 18
		glLineHeight = glFontSize * LINE_HEIGHT_RATIO
		
		Set gobjExplorer = CreateObject("InternetExplorer.Application")
		gobjExplorer.Navigate "about:blank"
		gobjExplorer.ToolBar = 0
		gobjExplorer.StatusBar = 0
		gobjExplorer.Width = 450
		gobjExplorer.Height = 200
		gobjExplorer.Left = ( glHorizontal - gobjExplorer.Width ) / 2
		gobjExplorer.Top = ( glVertical - gobjExplorer.Height ) / 2
		gobjExplorer.Visible = 1
		
		gobjExplorer.Document.Body.InnerHTML = ""
	End Sub
	
	Private Sub Class_Terminate()
		' Do Nothing
	End Sub
	
	Public Sub Activate()
		gobjExplorer.Document.Body.Style.Cursor = "wait" '�}�E�X�J�[�\���������v�ɂ���
		Call ActiveIE
	End Sub
	
	Public Sub Quit()
		gobjExplorer.Document.Body.Style.Cursor = "default" '�}�E�X�J�[�\�������ɖ߂�
		gobjExplorer.Quit
	End Sub
	
	'�E�B���h�E�̃T�C�Y�̓e�L�X�g�̍s���Ŏ����Z�o���邽�߁A�ݒ肳���Ȃ�
'	Public Property Let Height( _
'		ByVal lHeight _
'	)
'		gobjExplorer.Height = lHeight
'		gobjExplorer.Top = ( glVertical - gobjExplorer.Height ) / 2
'	End Property
	
	Public Property Let Width( _
		ByVal lWidth _
	)
		gobjExplorer.Width = lWidth
		gobjExplorer.Left = ( glHorizontal - gobjExplorer.Width ) / 2
	End Property
	
	Public Property Let Title( _
		ByVal sSetTitle _
	)
		gobjExplorer.Document.Title = sSetTitle
	End Property
	
	Public Property Let Font( _
		ByVal sFont _
	)
		gsFont = sFont
	End Property
	
	Public Property Let FontSize( _
		ByVal lFontSize _
	)
		glFontSize = lFontSize
		glLineHeight = lFontSize * LINE_HEIGHT_RATIO
	End Property
	
	Public Property Let Text( _
		ByVal sText _
	)
		sText = Replace( sText, vbNewLine, "<br>" )
		sText = Replace( sText, vbCr, "<br>" )
		sText = Replace( sText, vbLf, "<br>" )
		
		'�E�B���h�E�̍����A�ʒu�Z�o
		Dim lLineNum
		lLineNum = ( ( Len( sText ) - Len( Replace( sText, "<br>", "" ) ) ) / 4 ) + 1
		gobjExplorer.Height = ( ( glLineHeight * WIN_LINE_HEIGHT_RATIO ) * lLineNum ) + HEADER_HEIGHT
		gobjExplorer.Top = ( glVertical - gobjExplorer.Height ) / 2
		
	'	MsgBox lLineNum & "�F" & sText
		
		'�e�L�X�g�ݒ�
		gobjExplorer.Document.Body.InnerHTML = _
			"<font face=""" & gsFont & """>" & _
			"<span style=""font-size:" & glFontSize & "px; line-height:" & glLineHeight & "px;"">" & _
			sText & _
			"</span></font>"
	End Property
	
	Private Function ActiveIE()
		Dim Service
		Dim QfeSet
		Dim Qfe
		Dim lProcID
		Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
		Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='"& "iexplore.exe" &"'")
		lProcID = 0
		For Each Qfe in QfeSet
			lProcID = Qfe.ProcessId
			Exit For
		Next
		
		Dim objWshShell
		Set objWshShell = CreateObject("Wscript.Shell")
		objWshShell.AppActivate lProcID
		Set objWshShell = Nothing
	End Function
End Class
	If WScript.ScriptName = "IE.vbs" Then
		Call Test_IE
	End If
	Private Sub Test_IE
		Dim oIE
		Set oIE = New IE
		
		oIE.Activate
		WScript.Sleep(1000)
	'	Select Case 1
	'		Case 1:  oIE.Text = "��"
	'		Case 2:  oIE.Text = "��" & vbNewLine & "��"
	'		Case 3:  oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��"
	'		Case 4:  oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
	'		Case 8:  oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
	'		Case 16: oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
	'		Case Else: MsgBox "error!"
	'	End Select
	
		oIE.Title = "�^�C�g��"
		oIE.Font = "���C���I"
		oIE.FontSize = 30
		oIE.Text = "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
		WScript.Sleep(500)
		
		oIE.Font = "MS ����"
		oIE.FontSize = 8
		oIE.Text = "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
		WScript.Sleep(500)
		oIE.Text = "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��" & vbNewLine & "��"
		WScript.Sleep(500)
		
		oIE.Width = 500
		WScript.Sleep(1000)
		
		oIE.Quit
		Set oIE = Nothing
	End Sub
