Option Explicit

Dim objArgs
Set objArgs = WScript.Arguments

If objArgs.Count < 2 Then
    WScript.Echo "Usage: cscript ScpFetchFromRemote.vbs <hostname> <source_files_path>"
    WScript.Quit 1
End If

Dim objWshShell
Dim objFSO
Set objWshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sUserProfile
Dim sSshConfigPath
Dim sHostName
Dim sSourceFilesPath
sUserProfile = objWshShell.ExpandEnvironmentStrings("%USERPROFILE%")
sSshConfigPath = """" & sUserProfile & "\.ssh\config" & """"
sHostName = objArgs(0)
sSourceFilesPath = """" & sHostName & ":" & objArgs(1) & """"

Dim sDesktopPath
Dim sMyDirPathDesktop
Dim sTargetDirPath
sDesktopPath = objWshShell.ExpandEnvironmentStrings( "%MYDIRPATH_DESKTOP%" )
sMyDirPathDesktop = sDesktopPath
sTargetDirPath = """" & sMyDirPathDesktop & "\_scp" & """"

If Not objFSO.FolderExists(sMyDirPathDesktop & "\_scp") Then
    objFSO.CreateFolder sMyDirPathDesktop & "\_scp"
End If

Dim sCmd
sCmd = "scp -C -r -F " & sSshConfigPath & " " & sSourceFilesPath & " " & sTargetDirPath

objWshShell.Run sCmd, 1, True
