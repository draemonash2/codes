Option Explicit

Dim objArgs
Set objArgs = WScript.Arguments

If objArgs.Count < 3 Then
    WScript.Echo "Usage: cscript ScpUploadToRemote.vbs <hostname> <target_dir_path> <file_or_dir1> [file_or_dir2] ..."
    WScript.Quit 1
End If

Dim objWshShell
Dim objFSO
Set objWshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sUserProfile
Dim sSshConfigPath
Dim sHostName
Dim sTargetDirPath
sUserProfile = objWshShell.ExpandEnvironmentStrings("%USERPROFILE%")
sSshConfigPath = """" & sUserProfile & "\.ssh\config" & """"
sHostName = objArgs(0)
sTargetDirPath = """" & sHostName & ":" & objArgs(1) & """"

Dim i
For i = 2 To objArgs.Count - 1
    Dim sSourcePath
    Dim sCmd

    sSourcePath = objArgs(i)
    sCmd = ""
    If objFSO.FolderExists(sSourcePath) Then
        sCmd = "scp -r -F " & sSshConfigPath & " """ & sSourcePath & """ " & sTargetDirPath
    ElseIf objFSO.FileExists(sSourcePath) Then
        sCmd = "scp -F " & sSshConfigPath & " """ & sSourcePath & """ " & sTargetDirPath
    Else
        WScript.Echo "Warning: Not found - " & sSourcePath
        sCmd = ""
    End If

    If sCmd <> "" Then
        objWshShell.Run sCmd, 1, True
    End If
Next
