Option Explicit

Const sDownloadUrl = "https://github.com/draemonash2/codes/archive/master.zip"
Const sDownloadTrgtFileName = "codes.zip"
Const sDiffSrcDirName = "codes-master"
Const sDiffTrgtDirPath = "C:\codes"
Const lPopupWaitSecond = 5

Dim objWshShell
Set objWshShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sPopupTitle
sPopupTitle = WScript.ScriptName
Dim sDownloadTrgtDirPath
sDownloadTrgtDirPath = objWshShell.SpecialFolders("Desktop")
Dim sDownloadTrgtFilePath
sDownloadTrgtFilePath = sDownloadTrgtDirPath & "\" & sDownloadTrgtFileName
Dim sDiffSrcDirPath
sDiffSrcDirPath = sDownloadTrgtDirPath & "\" & sDiffSrcDirName

'=== �_�E�����[�h ===
Dim sPopupMsg
sPopupMsg = "�_�E�����[�h���J�n���܂��c"
objWshShell.Popup sPopupMsg, lPopupWaitSecond, sPopupTitle, vbInformation
Call DownloadFile( sDownloadUrl, sDownloadTrgtFilePath )

'=== �� ===
sPopupMsg = "�_�E�����[�h����!" & vbNewLine & "�𓀂��J�n���܂��c"
objWshShell.Popup sPopupMsg, lPopupWaitSecond, sPopupTitle, vbInformation
With CreateObject("Shell.Application")
    .NameSpace(sDownloadTrgtDirPath).CopyHere .NameSpace(sDownloadTrgtFilePath).Items
End With

'=== ��r ===
sPopupMsg = "�𓀊���!" & vbNewLine & "��r���J�n���܂��c"
objWshShell.Popup sPopupMsg, lPopupWaitSecond, sPopupTitle, vbInformation

Dim sDiffProgramPath
sDiffProgramPath = objWshShell.Environment("System").Item("MYSYSPATH_WINMERGE")
If sDiffProgramPath = "" then
	MsgBox "���ϐ����ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbYes, PROG_NAME
	WScript.Quit
end if
objWshShell.Run sDiffProgramPath & " " & sDiffSrcDirPath & " " & sDiffTrgtDirPath, 0, True

'[�Q�lURL] https://viewse.blogspot.com/2013/08/vbscriptweb.html
Private Function DownloadFile( _
    ByVal sDownloadUrl, _
    ByVal sDownloadTrgtFilePath _
)
    ' �_�E�����[�h�p�̃I�u�W�F�N�g
    Dim objSrvHTTP
    Set objSrvHTTP = Wscript.CreateObject("Msxml2.ServerXMLHTTP")
    on error resume next
    Call objSrvHTTP.Open("GET", sDownloadUrl, False )
    if Err.Number <> 0 then
        Wscript.Echo Err.Description
        Wscript.Quit
    end if
    objSrvHTTP.Send
    
    if Err.Number <> 0 then
    ' �����炭�T�[�o�[�̎w�肪�Ԉ���Ă���
        Wscript.Echo Err.Description
        Wscript.Quit
    end if
    on error goto 0
    if objSrvHTTP.status = 404 then
        Wscript.Echo "URL������������܂���(404)"
        Wscript.Quit
    end if
    
    ' �o�C�i���f�[�^�ۑ��p�I�u�W�F�N�g
    Dim Stream
    Set Stream = Wscript.CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = 1 ' �o�C�i��
    ' �߂��ꂽ�o�C�i�����t�@�C���Ƃ��ăX�g���[���ɏ�������
    Stream.Write objSrvHTTP.responseBody
    ' �t�@�C���Ƃ��ĕۑ�
    Stream.SaveToFile sDownloadTrgtFilePath, 2
    Stream.Close
End Function

