Dim sTime
sTime = Time
sTime = Left(sTime, InStrRev(sTime, ":") - 1)
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
objWshShell.Popup sTime & "�ł�", 10, sTime & "�ł�", vbInformation

