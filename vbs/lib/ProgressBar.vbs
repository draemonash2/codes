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
		
		goIE.Title = "�i����"
		goIE.Activate
		
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
			sProgMsg = gsProgMsg & String( 2, vbNewLine )
		End If
		sProgMsg = _
			sProgMsg & "������..."& vbNewLine & _
			DateDiff( "s", glStartTime, Now() ) & " [s] �o��..." & vbNewLine & _
			String( glProg10, "��") & String( 10 - glProg10, "��") & "  " & glProg100 & "% ����"
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
