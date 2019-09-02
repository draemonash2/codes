Dim sTime
sTime = Time
sTime = Left(sTime, InStrRev(sTime, ":") - 1)
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
objWshShell.Popup sTime & "‚Å‚·", 10, sTime & "‚Å‚·", vbInformation

