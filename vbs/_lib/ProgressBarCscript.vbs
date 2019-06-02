Option Explicit

Dim objPrgBar, i, lMaxValue
Set objPrgBar = New ProgressBar

'Cscript�œ������Ă��邩���肵�K�v�ɉ����ď�����ς���
Call objPrgBar.CscriptRun()

'#�����P
Call objPrgBar.ShowMessage("���������@���s!")
lMaxValue = 255
For i = 1 To lMaxValue
	WScript.Sleep 1
	Call objPrgBar.ShowProgress(i, lMaxValue)
Next

'#�����Q
Call objPrgBar.ShowMessage("�Z�������@���s!")
lMaxValue= 10
For i = 1 To lMaxValue
	WScript.Sleep 45
	Call objPrgBar.ShowProgress(i, lMaxValue)
Next

Call objPrgBar.ShowMessage("Complete!!")
msgbox "�I�����܂���"


Class ProgressBar
	Private sStatus
	Private sExecName
	Private iPercentage
	Private sPercentageStr
	Private sProgressBar
	Private sScriptPath
	Private objFSO
	Private objWshShell
	
	'=========================================
	'= �R���X�g���N�^
	'=========================================
	Private Sub Class_Initialize
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objWshShell = WScript.CreateObject("WScript.Shell")
		sScriptPath = Wscript.ScriptFullName
	End Sub
	
	Private Sub Class_Terminate
		Set objFSO = Nothing
		Set objWshShell = Nothing
	End Sub
	
	'=========================================
	'= �p�u���b�N�֐�
	'=========================================
	Public Sub CscriptRun()
		if IsCscript() then
			'Do Nothing
		else
			objWshShell.Run "cscript //nologo """ & sScriptPath & """", 1, False
			Wscript.Quit
		end if
	End Sub
	
	Public Sub ShowProgress( _
		ByVal lBunsi, _
		ByVal lBunbo _
	)
		Call MakePercentage(lBunsi, lBunbo)
		Call MakeProgressBar()
		Wscript.StdOut.Write sPercentageStr & " |" & sProgressBar & "| " & lBunsi & "/" & lBunbo & vbCr
		sStatus = "ShowProgress"
	End Sub
	
	Public Sub ShowMessage( _
		Byval sMessage _
	)
		if sStatus = "ShowProgress" then
			Wscript.StdOut.Write vbCrLf
		end if
		Wscript.StdOut.Write sMessage & vbCrLf
		sStatus = "ShowMessage"
	End SUb
	
	'=========================================
	'= �v���C�x�[�g�֐�
	'=========================================
	Private Function IsCscript()
		sExecName = LCase(objFSO.GetFileName(WScript.FullName))
		if sExecName = "cscript.exe" then
			IsCscript = true
		else
			IsCscript = false
		end if
	End Function
	
	Private Sub MakePercentage( _
		ByVal lBunsi, _
		ByVal lBunbo _
	)
		iPercentage = Cint((lBunsi / lBunbo) * 100)
		sPercentageStr = iPercentage & "%"
		sPercentageStr = String(4 - Len(sPercentageStr), " ") & sPercentageStr
	End Sub
	
	Private Sub MakeProgressBar()
		sProgressBar = String(Cint(iPercentage/5), "=") & ">" & String(20 - Cint(iPercentage/5), " ")
	End Sub
	
End Class
