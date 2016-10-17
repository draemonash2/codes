Option Explicit

'==========================================================
'= �ݒ�l
'==========================================================
Const TRGT_DIR = "Z:\300_Musics"
Const UPDATE_MOD_DATE = False

Const DEBUG_FUNCVALID_ADDFILES        = True
Const DEBUG_FUNCVALID_DATEINPUT       = True
Const DEBUG_FUNCVALID_TRGTLISTUP      = True
Const DEBUG_FUNCVALID_DIRCMDEXEC      = True
Const DEBUG_FUNCVALID_TAGUPDATE       = True
Const DEBUG_FUNCVALID_DIRRESULTDELETE = True

'==========================================================
'= �{����
'==========================================================
Dim objWshShell
Dim sCurDir
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
sCurDir = objWshShell.CurrentDirectory
Call Include( sCurDir & "\lib\String.vbs" )
Call Include( sCurDir & "\lib\StopWatch.vbs" )
Call Include( sCurDir & "\lib\ProgressBar.vbs" )
Call Include( sCurDir & "\lib\FileSystem.vbs" )
Call Include( sCurDir & "\lib\iTunes.vbs" )
Call Include( sCurDir & "\lib\Array.vbs" )

Dim sLogFilePath
sLogFilePath = TRGT_DIR & "\" & RemoveTailWord( WScript.ScriptName, "." ) & ".log"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objLogFile
Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )

objLogFile.WriteLine "[�X�V�Ώۃt�H���_] " & TRGT_DIR
objLogFile.WriteLine "[�X�V���� �ύX�L��] " & UPDATE_MOD_DATE

Dim oStpWtch
Set oStpWtch = New StopWatch
Call oStpWtch.StartT

Dim oPrgBar
Set oPrgBar = New ProgressBar

' ******************************************
' * �t�@�C���ǉ�                           *
' ******************************************
If DEBUG_FUNCVALID_ADDFILES Then ' ��Debug��

objLogFile.WriteLine "*** �t�@�C���ǉ� *** "
oPrgBar.SetMsg( _
	"�ˁE�t�@�C���ǉ�����" & vbNewLine & _
	"�@�E���t���͏���" & vbNewLine & _
	"�@�E�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�@�E�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 50 ) '�i���X�V

Dim sAnswer
sAnswer = MsgBox( "iTunes �փt�@�C����ǉ����܂����H" & vbNewLine & _
				  "  [�ǉ��Ώۃt�H���_] " & TRGT_DIR _
				  , vbYesNoCancel _
				)
If sAnswer = vbYes Then
	MsgBox "iTunes �� " & TRGT_DIR & " ��ǉ����܂��B"
	WScript.CreateObject("iTunes.Application").LibraryPlaylist.AddFile( TRGT_DIR )
ElseIf sAnswer = vbNo Then
	MsgBox "iTunes �ւ̒ǉ����X�L�b�v���܂��B"
Else
	MsgBox "�v���O�����𒆒f���܂��B"
	Call Finish
	WScript.Quit
End If
oPrgBar.SetProg( 100 ) '�i���X�V

objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime & " [s]"
End If ' ��Debug��

' ******************************************
' * ���t����                               *
' ******************************************
If DEBUG_FUNCVALID_DATEINPUT Then ' ��Debug��

objLogFile.WriteLine ""
objLogFile.WriteLine "*** ���t���͏��� *** "
oPrgBar.SetMsg( _
	"�@�E�t�@�C���ǉ�����" & vbNewLine & _
	"�ˁE���t���͏���" & vbNewLine & _
	"�@�E�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�@�E�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '�i���X�V

On Error Resume Next

oPrgBar.SetProg( 10 ) '�i���X�V

Dim sNow
sNow = Now()
sNow = Left( sNow, Len( sNow ) - 2 ) & "00" '�b��00�ɂ���

Dim sCmpBaseTime
sCmpBaseTime = InputBox( _
					"�X�V�ΏۂƂ���t�@�C������肵�܂��B" & vbNewLine & _
					"�X�V�ΏۂƂ��鎞������͂��Ă��������B" & vbNewLine & _
					"" & vbNewLine & _
					"  [���͋K��] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
					"" & vbNewLine & _
					"�� ���t�݂̂��w�肵�����ꍇ�A�uYYYY/MM/DD 0:0:0�v�Ƃ��Ă��������B" _
					, "����" _
					, sNow _
				)

objLogFile.WriteLine "���͒l : " & sCmpBaseTime

Dim sTimeValue
Dim sDateValue
sTimeValue = TimeValue(sCmpBaseTime)
sDateValue = DateValue(sCmpBaseTime)

oPrgBar.SetProg( 50 ) '�i���X�V

'���t�`�F�b�N
If Err.Number <> 0 Then
	MsgBox "���t�̌`�����s���ł��I" & vbNewLine & _
	       "  [���͋K��] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
	       "  [���͒l] " & sCmpBaseTime
	MsgBox Err.Description
	MsgBox "�v���O�����𒆒f���܂��I"
	Err.Clear
	Call Finish
	WScript.Quit
Else
	'Do Nothing
End If
If DateDiff("s", sCmpBaseTime, Now() ) < 0  Then
	MsgBox "�����̓������w�肳��܂����I" & vbNewLine & _
	       "  [���͒l] " & sCmpBaseTime
	MsgBox "�v���O�����𒆒f���܂��I"
	Call Finish
	WScript.Quit
Else
	'Do Nothing
End If
On Error Goto 0 '�uOn Error Resume Next�v������

oPrgBar.SetProg( 100 ) '�i���X�V

objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime & " [s]"

Else ' ��Debug��
sCmpBaseTime = "2016/10/16 22:50:00"
End If ' ��Debug��

' ******************************************
' * �X�V�Ώۃt�@�C�����X�g�擾             *
' ******************************************
If DEBUG_FUNCVALID_TRGTLISTUP Then ' ��Debug��

objLogFile.WriteLine ""
objLogFile.WriteLine "*** �X�V�Ώۃt�@�C������ *** "
oPrgBar.SetMsg( _
	"�@�E�t�@�C���ǉ�����" & vbNewLine & _
	"�@�E���t���͏���" & vbNewLine & _
	"�ˁE�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�@�E�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '�i���X�V

On Error Resume Next

'*** Dir �R�}���h���s ***
Dim sTmpFilePath
Dim sExecCmd
sTmpFilePath = objWshShell.CurrentDirectory & "\" & replace( WScript.ScriptName, ".vbs", "_TrgtFileList.tmp" )
If DEBUG_FUNCVALID_DIRCMDEXEC Then ' ��Debug��
sExecCmd = "Dir """ & TRGT_DIR & """ /s /a:a-d > """ & sTmpFilePath & """"
With CreateObject("Wscript.Shell")	
	.Run "cmd /c" & sExecCmd, 7, True
End With
End If ' ��Debug��

'*** Dir �R�}���h���ʎ擾 ***
Dim objFile
Dim sTextAll
If Err.Number = 0 Then
	Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
	If Err.Number = 0 Then
		sTextAll = objFile.ReadAll
		sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '�����ɉ��s���t�^����Ă��܂����߁A�폜
		objFile.Close
	Else
		WScript.Echo "�t�@�C�����J���܂���: " & Err.Description
	End If
	Set objFile = Nothing	'�I�u�W�F�N�g�̔j��
Else
	WScript.Echo "�G���[ " & Err.Description
End If
On Error Goto 0

oPrgBar.SetProg( 20 ) '�i���X�V

'*** �X�V�������o ***
Dim oMatchResult
Dim sSearchPattern
Dim oRegExp
Dim sTargetStr
Set oRegExp = CreateObject("VBScript.RegExp")
sSearchPattern = "((\d{4}/\d{1,2}/\d{1,2})\s+(\d{1,2}:\d{1,2})\s+([0-9,]+)\s+(.+)\r)|(\s+(.*)\s�̃f�B���N�g��)"
sTargetStr = sTextAll
oRegExp.Pattern = sSearchPattern               '�����p�^�[����ݒ�
oRegExp.IgnoreCase = True                      '�啶���Ə���������ʂ��Ȃ�
oRegExp.Global = True                          '������S�̂�����
Set oMatchResult = oRegExp.Execute(sTargetStr) '�p�^�[���}�b�`���s

Dim sFileName
Dim sFilePath
Dim sFileSize
Dim sModDate
Dim sDirName
Dim iMatchIdx
Dim sExtName
Dim asTrgtFileList()
ReDim asTrgtFileList(-1)
Dim lMatchResultCount
sDirName = ""
objLogFile.WriteLine "[sFilePath]" & chr(9) & _
					 "[sDirName]"  & chr(9) & _
					 "[sFileName]" & chr(9) & _
					 "[sModDate]"  & chr(9) & _
					 "[sFileSize]" & chr(9) & _
					 "[sExtName]"
lMatchResultCount = oMatchResult.Count
For iMatchIdx = 0 To lMatchResultCount - 1
	'�i���X�V
'	oPrgBar.SetProg( _
'		oPrgBar.ConvProgRange( _
'			0, _
'			oMatchResult.Count - 1, _
'			iMatchIdx _
'		) _
'	)
	If oMatchResult(iMatchIdx).SubMatches.Count = 7 Then
		'�t�@�C�����Ƀ}�b�`
		If oMatchResult(iMatchIdx).SubMatches(0) <> "" Then
			sModDate = oMatchResult(iMatchIdx).SubMatches(1) & " " & _
			           oMatchResult(iMatchIdx).SubMatches(2)
			sFileSize = oMatchResult(iMatchIdx).SubMatches(3)
			sFileName = oMatchResult(iMatchIdx).SubMatches(4)
			sFilePath = sDirName & "\" & sFileName
			sExtName = ExtractTailWord( sFileName, "." )
			
'			objLogFile.WriteLine oMatchResult(iMatchIdx).SubMatches(0) & chr(9) & _
'			                     oMatchResult(iMatchIdx).SubMatches(1) & chr(9) & _
'			                     oMatchResult(iMatchIdx).SubMatches(2) & chr(9) & _
'			                     oMatchResult(iMatchIdx).SubMatches(3) & chr(9) & _
'			                     oMatchResult(iMatchIdx).SubMatches(4) & chr(9) & _
'			                     oMatchResult(iMatchIdx).SubMatches(5) & chr(9) & _
'			                     oMatchResult(iMatchIdx).SubMatches(6)
			
			'�X�V������r �� �X�V�ΏۑI��
			If LCase(sExtName) = "mp3" Then
				If DateDiff("s", sCmpBaseTime, sModDate ) >= 0  Then
					ReDim Preserve asTrgtFileList( UBound(asTrgtFileList) + 1 )
					asTrgtFileList( UBound(asTrgtFileList) ) = sFilePath
					objLogFile.WriteLine sFilePath & chr(9) & _
										 sDirName  & chr(9) & _
										 sFileName & chr(9) & _
										 sModDate  & chr(9) & _
										 sFileSize & chr(9) & _
										 sExtName
				Else
					'Do Nothing
				End If
			Else
				'Do Nothing
			End If
			
		'�t�H���_���Ƀ}�b�`
		ElseIf oMatchResult(iMatchIdx).SubMatches(5) <> "" Then
			sDirName = oMatchResult(iMatchIdx).SubMatches(6)
		Else
			MsgBox "�G���[�I"
		End If
	Else
		MsgBox "�G���[�I"
	End If
Next

oPrgBar.SetProg( 100 ) '�i���X�V

If DEBUG_FUNCVALID_DIRRESULTDELETE Then ' ��Debug��
objFSO.DeleteFile sTmpFilePath, True
End If ' ��Debug��

Set objFSO = Nothing	'�I�u�W�F�N�g�̔j��

objLogFile.WriteLine "�t�@�C�����F" & UBound(asTrgtFileList) + 1
objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime & " [s]"

Else ' ��Debug��
	ReDim asTrgtFileList(0)
	asTrgtFileList(0) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Bow Down.mp3"
'	asTrgtFileList(1) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concentrate.mp3"
'	asTrgtFileList(2) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concrete Schoolyard.mp3"
'	asTrgtFileList(3) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Control Myself.mp3"
End If ' ��Debug��

' ******************************************
' * �^�O�X�V                               *
' ******************************************
If DEBUG_FUNCVALID_TAGUPDATE Then ' ��Debug��
	
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �^�O�X�V���� *** "
oPrgBar.SetMsg( _
	"�@�E�t�@�C���ǉ�����" & vbNewLine & _
	"�@�E���t���͏���" & vbNewLine & _
	"�@�E�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�ˁE�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '�i���X�V

objLogFile.WriteLine "[FilePath]" & Chr(9) & "[TrackName}" & Chr(9) & "[HitNum]"

Dim lTrgtFileListIdx
Dim lTrgtFileListNum
lTrgtFileListNum = UBound( asTrgtFileList )
For lTrgtFileListIdx = 0 To lTrgtFileListNum
	'�i���X�V
	oPrgBar.SetProg( _
		oPrgBar.ConvProgRange( _
			0, _
			lTrgtFileListNum, _
			lTrgtFileListIdx _
		) _
	)
	
	Dim sTrgtFilePath
	sTrgtFilePath = asTrgtFileList( lTrgtFileListIdx )
	
	'�g���b�N���擾
	Dim sTrgtDirPath
	Dim sTrgtFileName
	sTrgtDirPath = RemoveTailWord( sTrgtFilePath, "\" )
	sTrgtFileName = ExtractTailWord( sTrgtFilePath, "\" )
	
	Dim oTrgtDirFiles
	Dim oTrgtFile
	Dim sTrgtTrackName
	Dim sTrgtModDate
	Set oTrgtDirFiles = CreateObject("Shell.Application").Namespace( sTrgtDirPath )
	Set oTrgtFile = oTrgtDirFiles.ParseName( sTrgtFileName )
	sTrgtTrackName = oTrgtDirFiles.GetDetailsOf( oTrgtFile, 21 )
	sTrgtModDate = oTrgtFile.ModifyDate
	
	Dim objPlayList
	Dim objSearchResult
	Set objPlayList = WScript.CreateObject("iTunes.Application").Sources.Item(1).Playlists.ItemByName("�~���[�W�b�N")
	Set objSearchResult = objPlayList.Search( sTrgtTrackName, 5 )
	
	objLogFile.WriteLine sTrgtFilePath & Chr(9) & sTrgtTrackName & Chr(9) & objSearchResult.Count
	
	Dim lHitIdx
	For lHitIdx = 1 to objSearchResult.Count
		With objSearchResult.Item(lHitIdx)
			If .Location = sTrgtFilePath Then
				.Composer = "1"
				.Composer = ""
				Exit For
			Else
				'Do Nothing
			End If
		End With
	Next
	Set objSearchResult = Nothing
	Set objPlayList = Nothing
	
	If UPDATE_MOD_DATE = True Then
		'Do Nothing
	Else
		oTrgtFile.ModifyDate = CDate( sTrgtModDate )
	End If
	
	Set oTrgtFile = Nothing
	Set oTrgtDirFiles = Nothing
Next

objLogFile.WriteLine "�t�@�C�����F" & UBound(asTrgtFileList) + 1
objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime & " [s]"

End If ' ��Debug��

' ******************************************
' * �I������                               *
' ******************************************
Call Finish
MsgBox "�v���O����������ɏI�����܂����B"

'==========================================================
'= �֐���`
'==========================================================
' �O���v���O���� �C���N���[�h�֐�
Function Include( _
	ByVal sOpenFile _
)
	Dim objFSO
	Dim objVbsFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
	
	ExecuteGlobal objVbsFile.ReadAll()
	objVbsFile.Close
	
	Set objVbsFile = Nothing
	Set objFSO = Nothing
End Function

Function Finish()
	Call oStpWtch.StopT
	Call oPrgBar.Quit
	objLogFile.WriteLine ""
	objLogFile.WriteLine "�J�n����               : " & oStpWtch.StartPoint
	objLogFile.WriteLine "�I������               : " & oStpWtch.StopPoint
	objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime & " [s]"
	objLogFile.Close
	Set oStpWtch = Nothing
	Set oPrgBar = Nothing
End Function