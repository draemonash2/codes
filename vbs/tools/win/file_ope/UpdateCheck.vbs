Option Explicit

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" ) 'DownloadFile()

'===============================================================================
'= �{����
'===============================================================================
Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count <> 3 Then
    MsgBox "�����𐳂����w�肵�Ă��������B", vbExclamation, WScript.ScriptName
    WScript.Quit
End If
Dim sDownloadUrl
Dim sLocalObjName
Dim sDiffTrgtDirPath
sDiffTrgtDirPath = WScript.Arguments(0)
sDownloadUrl = WScript.Arguments(1)
sLocalObjName = WScript.Arguments(2)

Dim sOutputMsg
sOutputMsg = WScript.ScriptName & " �u" & sLocalObjName & "�v"

'=== ���O���� ===
Dim sDownloadTrgtDirPath
Dim sDownloadTrgtFilePath
Dim sDiffSrcDirPath
Dim sUnzipProgramPath
Dim sDiffProgramPath
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sLocalObjName & ".zip"
sDiffSrcDirPath = sDownloadTrgtDirPath & "\" & sLocalObjName
sUnzipProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_7Z%")
If InStr(sUnzipProgramPath, "%") > 0 then
    MsgBox "���ϐ��uMYEXEPATH_7Z�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sOutputMsg
    WScript.Quit
End If
sDiffProgramPath = objWshShell.ExpandEnvironmentStrings("%MYEXEPATH_WINMERGE%")
If InStr(sDiffProgramPath, "%") > 0 then
    MsgBox "���ϐ��uMYEXEPATH_WINMERGE�v���ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbExclamation, sOutputMsg
    WScript.Quit
end if

Dim vAnswer
vAnswer = MsgBox("�_�E�����[�h���J�n���܂��B", vbOkCancel, sOutputMsg)
If vAnswer = vbCancel Then
    MsgBox "�L�����Z���������ꂽ���߁A�����𒆒f���܂��B", vbExclamation, sOutputMsg
    WScript.Quit
End If

'=== �_�E�����[�h ===
Call DownloadFile(sDownloadUrl, sDownloadTrgtFilePath)

'=== �� ===
objWshShell.Run """" & sUnzipProgramPath & """ x -o""" & sDownloadTrgtDirPath & """ """ & sDownloadTrgtFilePath & """", 0, True

'=== �t�H���_��r ===
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcDirPath & """ """ & sDiffTrgtDirPath & """", 10, True

'=== �t�H���_�폜 ===
vAnswer = MsgBox("�_�E�����[�h�t�H���_���폜���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objFSO.DeleteFile sDownloadTrgtFilePath, True
    objFSO.DeleteFolder sDiffSrcDirPath, True
End If

MsgBox "�������������܂����I", vbYesNo, sOutputMsg

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

