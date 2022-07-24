Option Explicit

'<<�T�v>>
'  �w�肵���t�H���_�z���̃t�@�C�����o�b�N�A�b�v����B
'  
'<<�g�p���@>>
'  BackUpFilesAll.vbs <scriptpath> <rootdirpath> <filepathpattern> <backupnum> <backuplogpath>
'  
'<<�d�l>>
'  �E<rootdirpath> �z���̒��� <filepathpattern> �Ƀ}�b�`����t�@�C�����o�b�N�A�b�v����B
'    �i_bak�t�H���_�z���̂��̂͑ΏۊO�j
'  �E�o�b�N�A�b�v�̎d�l�� <scriptpath> �ɏ�����B

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()

'===============================================================================
'= �ݒ�l
'===============================================================================

'===============================================================================
'= �{����
'===============================================================================
Const sSCRIPT_NAME = "�t�@�C���ꊇ�o�b�N�A�b�v"

Dim sBakScriptPath
Dim sBakSrcRootDirPath
Dim sExtractFileNamePattern
Dim lBakFileNum
Dim sBakSrcLogPath
If WScript.Arguments.Count >= 5 Then
    sBakScriptPath = WScript.Arguments(0)
    sBakSrcRootDirPath = WScript.Arguments(1)
    sExtractFileNamePattern = WScript.Arguments(2)
    lBakFileNum = CLng(WScript.Arguments(3))
    sBakSrcLogPath = WScript.Arguments(4)
Else
    WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
    WScript.Quit
End If

Dim cFilePaths
Set cFilePaths = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sBakSrcRootDirPath, cFilePaths, 1, "")

Dim oRegExp1
Set oRegExp1 = CreateObject("VBScript.RegExp")
Dim oRegExp2
Set oRegExp2 = CreateObject("VBScript.RegExp")
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

oRegExp1.Pattern = sExtractFileNamePattern & "$"
oRegExp1.IgnoreCase = True
oRegExp1.Global = True
oRegExp2.Pattern = "\\_bak\\" & sExtractFileNamePattern & "$"
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
