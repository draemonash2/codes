Option Explicit

Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim cFilePathList
Set cFilePathList = CreateObject("System.Collections.ArrayList")

If WScript.Arguments.Count <> 1 Then
    MsgBox "�����𐳂����w�肵�Ă��������B", vbExclamation, WScript.ScriptName
    WScript.Quit
End If
Dim sTrgtDirPath
sTrgtDirPath = WScript.Arguments(0)

Dim sDiffProgramPath
sDiffProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_WINMERGE%")
If InStr(sDiffProgramPath, "%") > 0 then
    MsgBox "���ϐ��uMYEXEPATH_WINMERGE�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, WScript.ScriptName
    WScript.Quit
End if

'�t�H���_��r�Ώۑ������t�H���_��r���s
Dim vDirPath
Dim sDirNameRaw
Dim sDirNameBase
Dim sParDirPath
Dim sDiffSrcDirPath
Dim sDiffTrgtDirPath
dim oFolder
set oFolder = objFSO.getFolder(sTrgtDirPath)
For each vDirPath in oFolder.subfolders
    'MsgBox vDirPath
    sParDirPath = objFSO.GetParentFolderName( vDirPath )
    sDirNameRaw = objFSO.GetFileName( vDirPath )
    If InStr(sDirNameRaw, "_local") > 0 Then
        sDirNameBase = Replace(sDirNameRaw, "_local", "")
        sDiffSrcDirPath = sParDirPath & "\" & sDirNameBase
        sDiffTrgtDirPath = sParDirPath & "\" & sDirNameRaw
        If objFSO.FolderExists( sParDirPath & "\" & sDirNameBase ) Then
            objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 10, False
        Else
            'Do Nothing
        End If
    Else
        'Do Nothing
    End If
Next

'MsgBox "�����I", vbOkOnly, WScript.ScriptName

