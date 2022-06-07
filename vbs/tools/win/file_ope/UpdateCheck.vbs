Option Explicit

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Url.vbs" )    'DownloadFile()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" ) 'ConvDate2String()

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
Dim sDiffSrcOrgDirPath
Dim sDiffSrcNewDirPath
Dim sUnzipProgramPath
Dim sDiffProgramPath
Dim sDateSuffix
sDateSuffix = ConvDate2String(Now(),1)
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sLocalObjName & ".zip"
sDiffSrcOrgDirPath = sDownloadTrgtDirPath & "\" & sLocalObjName
sDiffSrcNewDirPath = sDownloadTrgtDirPath & "\" & sLocalObjName & "_" & sDateSuffix
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
objFSO.MoveFolder sDiffSrcOrgDirPath, sDiffSrcNewDirPath

'=== �t�H���_��r ===
objWshShell.Run """" & sDiffProgramPath & """ -r """ & sDiffSrcNewDirPath & """ """ & sDiffTrgtDirPath & """", 10, True

'=== ZIP�A�t�H���_�폜 ===
objFSO.DeleteFile sDownloadTrgtFilePath, True
vAnswer = MsgBox("�_�E�����[�h�t�H���_���폜���܂����H", vbYesNo, sOutputMsg)
If vAnswer = vbYes Then
    objFSO.DeleteFolder sDiffSrcNewDirPath, True
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

