Option Explicit

Dim objWshShell
Dim objFileSys
Dim objFolder
Dim objFile
Dim sMyFileName
Dim sNewFileName
Dim sOldFileName

Set objWshShell = WScript.CreateObject( "WScript.Shell" )
Set objFileSys = CreateObject( "Scripting.FileSystemObject" )
Set objFolder = objFileSys.GetFolder( objWshShell.CurrentDirectory )

sMyFileName = WScript.ScriptName

For Each objFile In objFolder.Files
    sOldFileName = objFile.Name
    If sMyFileName = sOldFileName Then
        'Do Nothing
    Else
        sNewFileName = sOldFileName
        Do While InStr( sNewFileName, ".png.png" ) > 0
            sNewFileName = Replace( sNewFileName, ".png.png", ".png" )
        Loop
        If sNewFileName = sOldFileName Then
            'Do Nothing
        Else
            objFile.Name = sNewFileName
        End If
    End If
Next
 
Set objFolder  = Nothing
Set objFileSys = Nothing 
Set objWshShell = Nothing
