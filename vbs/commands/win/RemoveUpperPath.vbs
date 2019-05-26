Option Explicit

'usage
' cscript.exe .\RemoveUpperPath.vbs <src_path> <dir_name> <remeve_dir_level>
'
'usage ex.
' cscript.exe .\RemoveUpperPath.vbs "C:\test\900_�\�[�X�R�[�h\5A45V\for_PDC-0154\src_b\sid\sid_can_2E.c" "900_�\�[�X�R�[�h" 2
'   ��"src_b\sid\sid_can_2E.c"
' cscript.exe .\RemoveUpperPath.vbs "C:\test\900_�\�[�X�R�[�h\5A45V\for_PDC-0154\src_b\sid\sid_can_2E.c" "900_�\�[�X�R�[�h" 3 | clip
'   ��"sid\sid_can_2E.c" ���N���b�v�{�[�h�ɃR�s�[

'==================================================
'= �ݒ�
'==================================================
Const DEFAULT_MATCH_DIR_NAME = "codes"
Const DEFAULT_REMOVE_DIR_LEVEL = 2

'==================================================
'= �{����
'==================================================
Dim sInputPath
Dim sMatchDirName
Dim lRemeveDirLevel
sMatchDirName = DEFAULT_MATCH_DIR_NAME
lRemeveDirLevel = DEFAULT_REMOVE_DIR_LEVEL
If WScript.Arguments.Count = 1 Then
	sInputPath = WScript.Arguments(0)
ElseIf WScript.Arguments.Count = 2 Then
	sInputPath = WScript.Arguments(0)
	sMatchDirName = WScript.Arguments(1)
ElseIf WScript.Arguments.Count = 3 Then
	sInputPath = WScript.Arguments(0)
	sMatchDirName = WScript.Arguments(1)
	lRemeveDirLevel = WScript.Arguments(2)
Else
	WScript.StdOut.WriteLine "�w�肷��������قȂ�܂��F" & WScript.Arguments.Count
	WScript.StdOut.WriteLine "�����𒆒f���܂�"
	WScript.Quit
End If

Dim sRemoveDirLevelPath
sRemoveDirLevelPath = ""
Dim lIdx
For lIdx = 0 To lRemeveDirLevel - 1
	sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
Next

Dim sSearchPattern
Dim sTargetStr
sSearchPattern = ".*\\" & sMatchDirName & sRemoveDirLevelPath & "\\"
sTargetStr = sInputPath

Dim oRegExp
Set oRegExp = CreateObject("VBScript.RegExp")
oRegExp.Pattern = sSearchPattern				'�����p�^�[����ݒ�
oRegExp.IgnoreCase = True						'�啶���Ə���������ʂ��Ȃ�
oRegExp.Global = True							'������S�̂�����

Dim oMatchResult
Set oMatchResult = oRegExp.Execute(sTargetStr)	'�p�^�[���}�b�`���s

Dim sOutPath
sOutPath = ""
If oMatchResult.Count > 0 THen
	sOutPath = Replace( sInputPath, oMatchResult.item(0), "" )
Else
	sOutPath = sInputPath
End If

If sOutPath = "" Then
	'Do Nothing
Else
	Wscript.StdOut.WriteLine sOutPath
End If
