Const sADDIN_DIR_PATH = "C:\codes\vba\excel\AddIns"

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetFileListCmdClct()
Call Include( "C:\codes\vbs\_lib\Debug.vbs" )       'DebugPrintClct()

'===============================================================================
'= �{����
'===============================================================================
'�t�H���_�p�X�ꗗ�擾
Dim cDirPathList
Set cDirPathList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sADDIN_DIR_PATH, cDirPathList, 2, "")
'Call DebugPrintClct(cDirPathList) '���f�o�b�O�p��

'�R�s�[���t�H���_����
Dim sSrcDirPath
sSrcDirPath = ""
Dim vDirPathTmp
For Each vDirPathTmp In cDirPathList
    Dim oRegExp
    Dim sTargetStr
    Dim sSearchPattern
    Set oRegExp = CreateObject("VBScript.RegExp")
    sTargetStr = vDirPathTmp
    sSearchPattern = "\Tmp\d{8}$"
    oRegExp.Pattern = sSearchPattern
    oRegExp.IgnoreCase = True
    oRegExp.Global = True
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)
    If oMatchResult.Count > 0 Then
        sSrcDirPath = vDirPathTmp
        Exit For
    End If
Next
'MsgBox sSrcDirPath: WScript.Quit '���f�o�b�O�p��
If sSrcDirPath = "" Then
    Msgbox "�R�s�[���t�H���_��������Ȃ����߁A�����𒆒f���܂�", vbOkOnly, WScript.ScriptName
    WScript.Quit
End If

'�R�s�[���t�H���_���̃t�@�C�����X�g�擾
Dim cSrcFilePathList
Set cSrcFilePathList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sSrcDirPath, cSrcFilePathList, 1, "")
'Call DebugPrintClct(cSrcFilePathList) '���f�o�b�O�p��

'�R�s�[���t�H���_���̃t�@�C�����R�s�[��t�H���_�փR�s�[
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Dim sDebugMsg '���f�o�b�O�p��
'sDebugMsg = "" '���f�o�b�O�p��
Dim vSrcFilePath
Dim sDstDirPath
sDstDirPath = sADDIN_DIR_PATH & "\MyExcelAddin.bas\"
For Each vSrcFilePath In cSrcFilePathList
    objFSO.CopyFile vSrcFilePath, sDstDirPath
    'sDebugMsg = sDebugMsg & vbNewLine & vSrcFilePath & "��" & sDstDirPath '���f�o�b�O�p��
Next
'MsgBox sDebugMsg: WScript.Quit '���f�o�b�O�p��

'�R�s�[���t�H���_�폜
objFSO.DeleteFolder sSrcDirPath, True

Msgbox "�X�V�����I", vbOkOnly, WScript.ScriptName

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function


