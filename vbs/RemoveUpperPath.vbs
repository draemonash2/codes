Option Explicit

'==================================================
' = ��j
' =	  �y���́z
' =		�N���b�v�{�[�h�l�F"C:\t_endo\900_�\�[�X�R�[�h\5A45V\for_PDC-0154\src_b\sid\sid_can_2E.c"
' =		SOURCE_DIR_NAME�F"900_�\�[�X�R�[�h"
' =		REMOVE_DIR_LEVEL�F2
' =   �y�o�́z
' =		�N���b�v�{�[�h�l�F"src_b\sid\sid_can_2E.c"
'==================================================

'==================================================
'= �ݒ�
'==================================================
Const SOURCE_DIR_NAME = "codes"
Const REMOVE_DIR_LEVEL = 1

'==================================================
'= �{����
'==================================================
Dim sInputPath
Dim objHTML
Set objHTML = CreateObject("htmlfile")
'sInputPath = Trim(objHTML.ParentWindow.ClipboardData.GetData("text"))
sInputPath = "C:\codes\vbs\test.vbs"

'msgbox sInputPath

Dim sRemoveDirLevelPath
sRemoveDirLevelPath = ""
Dim lIdx
For lIdx = 0 To REMOVE_DIR_LEVEL - 1
	sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
Next

Dim sSearchPattern
Dim sTargetStr
sSearchPattern = ".*\\" & SOURCE_DIR_NAME & sRemoveDirLevelPath & "\\"
'Msgbox sSearchPattern
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

'msgbox sOutPath

'With CreateObject("Wscript.Shell").Exec("clip")
'  .StdIn.Write sOutPath
'End With
Wscript.StdOut.WriteLine sOutPath

test
