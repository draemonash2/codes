'�t�@�C���𕪊�����
'
'usage
' cscript.exe .\_split_text_file.vbs <input_file_path> <keyword> [<output_file_path1> <output_file_path2>]

Option Explicit

'===============================================================================
'= �C���N���[�h
'===============================================================================

'===============================================================================
'= �ݒ�l
'===============================================================================

'===============================================================================
'= �{����
'===============================================================================
Const sSCRIPT_NAME = "�t�@�C������"

Dim sInFilePath
Dim sKeyword
Dim sOutFilePath1
Dim sOutFilePath2
If WScript.Arguments.Count = 4 Then
    sInFilePath = WScript.Arguments(0)
    sKeyword = WScript.Arguments(1)
    sOutFilePath1 = WScript.Arguments(2)
    sOutFilePath2 = WScript.Arguments(3)
ElseIf WScript.Arguments.Count = 2 Then
    sInFilePath = WScript.Arguments(0)
    sKeyword = WScript.Arguments(1)
    sOutFilePath1 = sInFilePath & ".split01"
    sOutFilePath2 = sInFilePath & ".split02"
Else
    WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
    WScript.Quit
End If

Dim adoInStrm
Dim adoOut1Strm
Dim adoOut2Strm
Set adoInStrm = CreateObject("ADODB.Stream")
Set adoOut1Strm = CreateObject("ADODB.Stream")
Set adoOut2Strm = CreateObject("ADODB.Stream")
Call SetOpenFileInfo(adoInStrm)
Call SetOpenFileInfo(adoOut1Strm)
Call SetOpenFileInfo(adoOut2Strm)
adoInStrm.LoadFromFile sInFilePath

Dim lSplitLineIdx
lSplitLineIdx = 0
Dim vFile
Dim vLine
vFile = Split(adoInStrm.ReadText(-1), vbLf)
For Each vLine In vFile
    If InStr(vLine, sKeyword) Then
        Exit For
    End If
    lSplitLineIdx = lSplitLineIdx + 1
Next

Dim lLineIdx
lLineIdx = 0
For Each vLine In vFile
    If lLineIdx < lSplitLineIdx - 1 Then
        adoOut1Strm.WriteText vLine, 1
    Else
        adoOut2Strm.WriteText vLine, 1
    End If
    lLineIdx = lLineIdx + 1
Next

adoInStrm.Close
Call SaveNoBomFile(adoOut1Strm, sOutFilePath1)
Call SaveNoBomFile(adoOut2Strm, sOutFilePath2)

Private Function SetOpenFileInfo( ByRef adoStrm )
    With adoStrm
        .Type = 2
        .Charset = "UTF-8"
        .LineSeparator = 10
        .Open
    End With
End Function

Private Function SaveNoBomFile( ByRef adoOutStrm, ByVal sOutFilePath )
    With adoOutStrm
        .Position = 0 '�X�g���[���̈ʒu��0�ɂ���
        .Type = 1 '�f�[�^�̎�ނ��o�C�i���f�[�^�ɕύX
        .Position = 3 '�X�g���[���̈ʒu��3�ɂ���
        
        Dim byteData
        byteData = .Read '�X�g���[���̓��e���ꎞ�i�[�p�ϐ��ɕۑ�
        .Close '��U�X�g���[�������i���Z�b�g�j
        
        .Open '�X�g���[�����J��
        .Write byteData '�X�g���[���Ɉꎞ�i�[�����f�[�^�𗬂�����
        .SaveToFile sOutFilePath, 2
        .Close
    End With
End Function

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
