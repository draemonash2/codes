'�������ꂽ�o�C�i���t�@�C������������B
'�����������o�C�i���t�@�C����I������
'�h���b�O�A���h�h���b�v���邱�ƂŁA
'���̃X�N���v�g�Ɠ����t�H���_�ɁuJoinFile�v���쐬����B
'zip �t�@�C���̏ꍇ�A�t�@�C�������uJoinFile�v�� �uJoinFile.zip�v
'�ɕύX���Ă���𓀂��邱�ƁB

Option Explicit

Dim asWSArgs
Set asWSArgs = WScript.Arguments

Dim sFileStr
sFileStr = ""
If asWSArgs.Count <= 1 Then
    WScript.Echo "More Arguments!"
    WScript.Quit
Else
    sFileStr = asWSArgs( 0 )
    Dim lArgIdx
    For lArgIdx = 1 to ( asWSArgs.Count - 1 )
        sFileStr = sFileStr & " + " & asWSArgs( lArgIdx )
    Next
End If

Dim oFileSys
Dim sParentDirName
Dim sOutFileName
Set oFileSys = CreateObject("Scripting.FileSystemObject")
sParentDirName = oFileSys.GetParentFolderName(asWSArgs(0))
sOutFileName = "JoinFile"

Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c copy /b " & sFileStr & " " & sParentDirName & "\" & sOutFileName, 0, false
Set objShell = Nothing

