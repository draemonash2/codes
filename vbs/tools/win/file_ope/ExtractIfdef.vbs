Option Explicit

'<�T�v>
'  �\�[�X�R�[�h���́u#ifdef __ve__�v�����ƂɁA���ɍ��킹�������𒊏o���ďo�͂���B
'
'    ex. �utest.f�v��{�X�N���v�g�Ƀh���b�O���h���b�v�����ꍇ
'      "#ifdef __ve__" �`"#else" ���𒊏o�����\�[�X�R�[�h���utest_ifside.f�v�Ƃ��ďo�͂���B
'      "#else" �`"#endif" ���𒊏o�����\�[�X�R�[�h���utest_elseside.f�v�Ƃ��ďo�͂���B
'
'<�g����>
'  1. �����������\�[�X�R�[�h�t�@�C����{�X�N���v�g�Ƀh���b�O���h���b�v����B

Const sSCRIPT_NAME = "IFDEF���o�c�[��"

' ������ �ݒ� �������� ������
Const sKEYWORD_IF = "#ifdef __ve__"
Const sKEYWORD_ELSE = "#else /** __ve__ **/"
Const sKEYWORD_ENDIF = "#endif /** __ve__ **/"
Const sFILE_SUFFIX_IFSIDE = "ifside"
Const sFILE_SUFFIX_ELSESIDE = "elseside"
' ������ �ݒ� �����܂� ������

If WScript.Arguments.Count = 0 Then
	MsgBox "�Œ��͈������K�v�ł��B�����𒆒f���܂��B"
	WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sInputFilePath
Dim sOutIfSideFilePath
Dim sOutElseSideFilePath
sInputFilePath = WScript.Arguments(0)
sOutIfSideFilePath = _
	objFSO.GetParentFolderName( sInputFilePath ) & "\" & _
	objFSO.GetBaseName( sInputFilePath ) & _
	"_" & sFILE_SUFFIX_IFSIDE & _
	"." & objFSO.GetExtensionName( sInputFilePath )
sOutElseSideFilePath = _
	objFSO.GetParentFolderName( sInputFilePath ) & "\" & _
	objFSO.GetBaseName( sInputFilePath ) & _
	"_" & sFILE_SUFFIX_ELSESIDE & _
	"." & objFSO.GetExtensionName( sInputFilePath )

Dim adoInFile
Set adoInFile = CreateObject("ADODB.Stream")
Dim adoOutIfSideFile
Set adoOutIfSideFile = CreateObject("ADODB.Stream")
Dim adoOutElseSideFile
Set adoOutElseSideFile = CreateObject("ADODB.Stream")

adoInFile.Type = 2
adoInFile.Charset = "Shift_JIS"
adoInFile.LineSeparator = 10
adoInFile.Open
adoInFile.LoadFromFile sInputFilePath
adoOutIfSideFile.Type = 2
adoOutIfSideFile.Charset = "Shift_JIS"
adoOutIfSideFile.LineSeparator = 10
adoOutIfSideFile.Open
adoOutElseSideFile.Type = 2
adoOutElseSideFile.Charset = "Shift_JIS"
adoOutElseSideFile.LineSeparator = 10
adoOutElseSideFile.Open

Const lMODE_DEFAULT = 0
Const lMODE_IFDEF = 1
Const lMODE_ELSE = 2

Dim lLineNo
Dim lKeywordMode
Dim bIsError
Dim sErrMsg
lLineNo = 1
lKeywordMode = lMODE_DEFAULT
sErrMsg = ""

Do Until adoInFile.EOS
	Dim sLine
    sLine = adoInFile.ReadText(-2)
	If InStr(sLine, sKEYWORD_IF) > 0 Then
		If lKeywordMode = lMODE_IFDEF Then
			sErrMsg = sErrMsg & "#ifdef���#ifdef��������܂��� : "& lLineNo & "�s��" & vbNewLine
		ElseIf lKeywordMode = lMODE_ELSE Then
			sErrMsg = sErrMsg & "#else���#ifdef��������܂��� : "& lLineNo & "�s��" & vbNewLine
		Else
			'Do Nothing
		End If
		lKeywordMode = lMODE_IFDEF
	ElseIf InStr(sLine, sKEYWORD_ELSE) > 0 Then
		If lKeywordMode = lMODE_ELSE Then
			sErrMsg = sErrMsg & "#else���#else��������܂��� : "& lLineNo & "�s��" & vbNewLine
		ElseIf lKeywordMode = lMODE_DEFAULT Then
			sErrMsg = sErrMsg & "#endif���#else��������܂��� : "& lLineNo & "�s��" & vbNewLine
		Else
			'Do Nothing
		End If
		lKeywordMode = lMODE_ELSE
	ElseIf InStr(sLine, sKEYWORD_ENDIF) > 0 Then
		If lKeywordMode = lMODE_DEFAULT Then
			sErrMsg = sErrMsg & "#endif���#endif��������܂��� : "& lLineNo & "�s��" & vbNewLine
		Else
			'Do Nothing
		End If
		lKeywordMode = lMODE_DEFAULT
	Else
		'�s�o��
		If lKeywordMode = lMODE_IFDEF Then
			adoOutIfSideFile.WriteText sLine, 1
		ElseIf lKeywordMode = lMODE_ELSE Then
			adoOutElseSideFile.WriteText sLine, 1
		ElseIf lKeywordMode = lMODE_DEFAULT Then
			adoOutIfSideFile.WriteText sLine, 1
			adoOutElseSideFile.WriteText sLine, 1
		Else
			MsgBox "�\�����Ȃ��G���[�I lKeywordMode = " & lKeywordMode
		End If
	End If
	lLineNo = lLineNo + 1
Loop

If sErrMsg <> "" Then
	MsgBox _
		"�G���[��������܂����B" & vbNewLine & _
		"�����𒆒f���܂��B" & vbNewLine & _
		"---" & vbNewLine & _
		sErrMsg, _
		vbOkOnly, _
		sSCRIPT_NAME
Else
	adoOutIfSideFile.SaveToFile sOutIfSideFilePath, 2
	adoOutElseSideFile.SaveToFile sOutElseSideFilePath, 2
	MsgBox "�����I", vbOkOnly, sSCRIPT_NAME
End If

adoInFile.Close
adoOutIfSideFile.Close
adoOutElseSideFile.Close

