' システムファイル
Key="HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShowSuperHidden"
Set wShell=CreateObject("WScript.Shell")
wShell.RegWrite Key,0,"REG_DWORD"
Set ShellWindows=CreateObject("Shell.Application").Windows()
For Each ie In ShellWindows
   If TypeName(ie.Document)<>"HTMLDocument" Then ie.Refresh
Next
ShellWindows.Item.Refresh

' 隠しファイル
Key = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden"
set wShell = CreateObject("WScript.Shell")
wShell.RegWrite Key, 2, "REG_DWORD"
set wShell = nothing
