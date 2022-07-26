Option Explicit

'<<�T�v>>
'  �w�肵���t�H���_�z���̃t�@�C�����o�b�N�A�b�v����B
'  
'<<�g�p���@>>
'  BackUpMemoFiles.vbs <rootdirpath> <backupnum> <backuplogpath>
'  
'<<�d�l>>
'  �E<rootdirpath> �z���̒��� sEXTRACT_FILE_NAME_PATTERN �Ƀ}�b�`����t�@�C�����o�b�N�A�b�v����B
'    �i_bak�t�H���_�z���̂��̂͑ΏۊO�j
'  �E�o�b�N�A�b�v�̎d�l�� <scriptpath> �ɏ�����B
'  
'<<�ˑ��X�N���v�g>>
'  �EBackUpMemoFiles.vbs

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()

'===============================================================================
'= �ݒ�l
'===============================================================================
Const sEXTRACT_FILE_NAME_PATTERN = "\\#memo.*\.xlsm$"
Const sBACKUP_SCRIPT_NAME = "BackUpFile.vbs"

'===============================================================================
'= �{����
'===============================================================================
Const sSCRIPT_NAME = "�t�@�C���ꊇ�o�b�N�A�b�v"

Dim sBakSrcRootDirPath
Dim lBakFileNum
Dim sBakSrcLogPath
If WScript.Arguments.Count >= 3 Then
    sBakSrcRootDirPath = WScript.Arguments(0)
    lBakFileNum = CLng(WScript.Arguments(1))
    sBakSrcLogPath = WScript.Arguments(2)
Else
    WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
    WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim sBakScriptPath
sBakScriptPath = objFSO.GetParentFolderName( WScript.ScriptFullName ) & "\" & sBACKUP_SCRIPT_NAME

Dim cFilePaths
Set cFilePaths = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sBakSrcRootDirPath, cFilePaths, 1, "")

Dim oRegExp1
Set oRegExp1 = CreateObject("VBScript.RegExp")
Dim oRegExp2
Set oRegExp2 = CreateObject("VBScript.RegExp")
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

oRegExp1.Pattern = sEXTRACT_FILE_NAME_PATTERN
oRegExp1.IgnoreCase = True
oRegExp1.Global = True
oRegExp2.Pattern = "\\_bak" & sEXTRACT_FILE_NAME_PATTERN
oRegExp2.IgnoreCase = True
oRegExp2.Global = True

Dim oMatchResult
Dim vFilePath
For Each vFilePath In cFilePaths
    Set oMatchResult = oRegExp1.Execute(vFilePath)
    If oMatchResult.Count > 0 Then
        Set oMatchResult = oRegExp2.Execute(vFilePath)
        If oMatchResult.Count = 0 Then
            Dim sCmdStr
            sCmdStr = """" & sBakScriptPath & """ """ & vFilePath & """ " & lBakFileNum & " """ & sBakSrcLogPath & """"
            'WScript.Echo sCmdStr
            objWshShell.Run sCmdStr, 0, True
        End If
    End If
Next

'WScript.Echo "�o�b�N�A�b�v�����I", vbOKOnly, sSCRIPT_NAME

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
