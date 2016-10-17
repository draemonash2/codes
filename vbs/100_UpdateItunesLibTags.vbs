Option Explicit

'==========================================================
'= �ݒ�l
'==========================================================
Const TRGT_DIR = "Z:\300_Musics"
Const UPDATE_MOD_DATE = False

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
sLogFilePath = sCurDir & "\" & RemoveTailWord( WScript.ScriptName, "." ) & ".log"

Dim objLogFile
Set objLogFile = CreateObject("Scripting.FileSystemObject").OpenTextFile( sLogFilePath, 2, True )

objLogFile.WriteLine "[�X�V�Ώۃt�H���_] " & TRGT_DIR
objLogFile.WriteLine "[�X�V���� �ύX�L��] " & UPDATE_MOD_DATE

Dim oStpWtch
Set oStpWtch = New StopWatch
Call oStpWtch.StartT

Dim oPrgBar
Set oPrgBar = New ProgressBar

If 1 Then ' ��Debug��

' ******************************************
' * ���t����                               *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** ���t���͏��� *** "
oPrgBar.SetMsg( _
	"�ˁE���t���͏���" & vbNewLine & _
	"�@�E�S�t�@�C�����X�g�擾����" & vbNewLine & _
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

' ******************************************
' * �S�t�@�C�����X�g�擾                   *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �S�t�@�C�����X�g�擾 *** "
oPrgBar.SetMsg( _
	"�@�E���t���͏���" & vbNewLine & _
	"�ˁE�S�t�@�C�����X�g�擾����" & vbNewLine & _
	"�@�E�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�@�E�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 20 ) '�i���X�V

Dim asAllFileList 'GetFileList2() �̐����ɂ��o���A���g�^�Œ�`�B�o���A���g�^�z��Ƃ��ĕԋp�����B
Call GetFileList2(TRGT_DIR, asAllFileList, 1)
'Call OutputAllElement( asAllFileList ) ' ��Debug��

oPrgBar.SetProg( 100 ) '�i���X�V

objLogFile.WriteLine "�t�@�C�����F" & UBound(asAllFileList) + 1
objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime & " [s]"
objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime & " [s]"

' ******************************************
' * �X�V�Ώۃt�@�C������                   *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �X�V�Ώۃt�@�C�����菈�� *** "
oPrgBar.SetMsg( _
	"�@�E���t���͏���" & vbNewLine & _
	"�@�E�S�t�@�C�����X�g�擾����" & vbNewLine & _
	"�ˁE�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�@�E�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '�i���X�V

Dim asTrgtFileList()
ReDim asTrgtFileList(-1)

Dim oFileSys
Set oFileSys = CreateObject("Scripting.FileSystemObject")

Dim sFilePath
Dim sLastModDate
Dim sExtName
Dim lAllFileListIdx
For lAllFileListIdx = 0 to UBound(asAllFileList)
	'�i���X�V
	oPrgBar.SetProg( _
		oPrgBar.ConvProgRange( _
			0, _
			UBound(asAllFileList), _
			lAllFileListIdx _
		) _
	)
	
	sExtName = ExtractTailWord( asAllFileList(lAllFileListIdx), "." )
	sFilePath = asAllFileList(lAllFileListIdx)
	If LCase(sExtName) = "mp3" Then
		sLastModDate = oFileSys.GetFile(sFilePath).DateLastModified
		If DateDiff("s", sCmpBaseTime, sLastModDate ) >= 0  Then
			ReDim Preserve asTrgtFileList( UBound(asTrgtFileList) + 1 )
			asTrgtFileList( UBound(asTrgtFileList) ) = sFilePath
			objLogFile.WriteLine sFilePath
		Else
			'Do Nothing
		End If
	Else
		'Do Nothing
	End If
Next

Set oFileSys = Nothing

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
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �^�O�X�V���� *** "
oPrgBar.SetMsg( _
	"�@�E���t���͏���" & vbNewLine & _
	"�@�E�S�t�@�C�����X�g�擾����" & vbNewLine & _
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
