'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################
Const INCLUDE_DOUBLE_QUOTATION = False

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�t�@�C���p�X���R�s�["

Dim bIsContinue
bIsContinue = True

Dim cFilePaths

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
	If EXECUTION_MODE = 0 Then 'Explorer������s
		Set cFilePaths = CreateObject("System.Collections.ArrayList")
		Dim sArg
		For Each sArg In WScript.Arguments
			cFilePaths.add sArg
		Next
	ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
		Set cFilePaths = WScript.Col( WScript.Env("Selected") )
	Else '�f�o�b�O���s
		MsgBox "�f�o�b�O���[�h�ł��B"
		Set cFilePaths = CreateObject("System.Collections.ArrayList")
		cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
		cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
	End If
Else
	'Do Nothing
End If

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
	If cFilePaths.Count = 0 Then
		MsgBox "�t�@�C�����I������Ă��܂���", vbYes, PROG_NAME
		MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
		bIsContinue = False
	Else
		'Do Nothing
	End If
Else
	'Do Nothing
End If

'*** ���΃p�X�ɒu�� ***
Dim objWshShellEnv
Set objWshShellEnv = WScript.CreateObject("WScript.Shell").Environment("Process")
sMatchDirName = objWshShellEnv.Item("MATCH_DIR_NAME")
lRemeveDirLevel = objWshShellEnv.Item("REMOVE_DIR_LEVEL")
Dim oFilePath
Dim sRelativePath
If sMatchDirName = "" And lRemeveDirLevel = "" Then
	'Do Nothing
Else
	Dim cRltvFilePaths
	Set cRltvFilePaths = CreateObject("System.Collections.ArrayList")
	For Each oFilePath In cFilePaths
		cRltvFilePaths.Add ReplaceRelatevePath(oFilePath, sMatchDirName, lRemeveDirLevel)
	Next
	Set cFilePaths = cRltvFilePaths
End If

'*** �N���b�v�{�[�h�փR�s�[ ***
If bIsContinue = True Then
	Dim sOutString
	Dim bFirstStore
	bFirstStore = True
	For Each oFilePath In cFilePaths
		If bFirstStore = True Then
			sOutString = oFilePath
			bFirstStore = False
		Else
			sOutString = sOutString & vbNewLine & oFilePath
		End If
	Next
	CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
	'Do Nothing
End If

'���΃p�X�֒u��
Private Function ReplaceRelatevePath( _
	ByVal sInFilePath, _
	ByVal sMatchDirName, _
	ByVal lRemeveDirLevel _
)
	Dim sRemoveDirLevelPath
	sRemoveDirLevelPath = ""
	Dim lIdx
	For lIdx = 0 To lRemeveDirLevel - 1
		sRemoveDirLevelPath = sRemoveDirLevelPath & "\\.+?"
	Next
	
	Dim sSearchPattern
	Dim sTargetStr
	sSearchPattern = ".*\\" & sMatchDirName & sRemoveDirLevelPath & "\\"
	sTargetStr = sInFilePath
	
	Dim oRegExp
	Set oRegExp = CreateObject("VBScript.RegExp")
	oRegExp.Pattern = sSearchPattern				'�����p�^�[����ݒ�
	oRegExp.IgnoreCase = True						'�啶���Ə���������ʂ��Ȃ�
	oRegExp.Global = True							'������S�̂�����
	
	Dim oMatchResult
	Set oMatchResult = oRegExp.Execute(sTargetStr)	'�p�^�[���}�b�`���s
	
	Dim sOutFilePath
	sOutFilePath = ""
	If oMatchResult.Count > 0 THen
		sOutFilePath = Replace( sInFilePath, oMatchResult.item(0), "" )
	Else
		sOutFilePath = sInFilePath
	End If
	
	ReplaceRelatevePath = sOutFilePath
End Function
