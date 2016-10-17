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

Dim objItunes
Dim objTracks
Dim oPrgBar
Dim oStpWtch

Set objItunes = WScript.CreateObject("iTunes.Application")
Set objTracks = objItunes.LibraryPlaylist.Tracks
Set oPrgBar = New ProgressBar
Set oStpWtch = New StopWatch

Dim objLogFile
Dim sLogFilePath
Dim sDbFilePath
Dim objFso
Set objFso = CreateObject("Scripting.FileSystemObject")
sLogFilePath = sCurDir & "\" & RemoveTailWord( WScript.ScriptName, "." ) & ".log"
sDbFilePath  = sCurDir & "\" & RemoveTailWord( WScript.ScriptName, "." ) & ".db"

Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )

Call oStpWtch.StartT

' ******************************************
' * ���t����                               *
' ******************************************
objLogFile.WriteLine "*** ���t���͏��� *** "
oPrgBar.SetMsg( _
	"�ˁE���t���͏���" & vbNewLine & _
	"�@�E�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�@�EPersistent ID ���X�g�i�A�z�z��j�擾" & vbNewLine & _
	"�@�E�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '�i���X�V

On Error Resume Next
Dim sTimeValue
Dim sDateValue
Dim sCmpBaseTime

oPrgBar.SetProg( 10 ) '�i���X�V

sCmpBaseTime = InputBox( _
					"�X�V�ΏۂƂ���t�@�C������肵�܂��B" & vbNewLine & _
					"�X�V�ΏۂƂ��鎞������͂��Ă��������B" & vbNewLine & _
					"" & vbNewLine & _
					"  [���͋K��] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
					"" & vbNewLine & _
					"�� ���t�݂̂��w�肵�����ꍇ�A�uYYYY/MM/DD 0:0:0�v�Ƃ��Ă��������B" _
					, "����" _
					, Now() _
				)

objLogFile.WriteLine "[���͒l]   " & sCmpBaseTime

sTimeValue = TimeValue(sCmpBaseTime)
sDateValue = DateValue(sCmpBaseTime)

oPrgBar.SetProg( 50 ) '�i���X�V

If Err.Number <> 0 Then
	Err.Clear '�G���[�����N���A����
	MsgBox "���t�̌`�����s���ł��I" & vbNewLine & _
	       "  [���͋K��] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
	       "  [���͒l] " & sCmpBaseTime
	MsgBox "�v���O�����𒆒f���܂��I"
	Call ExecuteFinish
	WScript.Quit
End If
On Error Goto 0 '�uOn Error Resume Next�v������

oPrgBar.SetProg( 100 ) '�i���X�V

objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime

' ******************************************
' * �X�V�Ώۃt�@�C������                   *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �X�V�Ώۃt�@�C�����菈�� *** "
oPrgBar.SetMsg( _
	"�@�E���t���͏���" & vbNewLine & _
	"�ˁE�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�@�EPersistent ID ���X�g�i�A�z�z��j�擾" & vbNewLine & _
	"�@�E�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '�i���X�V

Dim asAllFileList()
Dim asTrgtFileList()
ReDim asAllFileList(-1)
ReDim asTrgtFileList(-1)

oPrgBar.SetProg( 20 ) '�i���X�V
Call GetFileList(TRGT_DIR, asAllFileList, 1)
oPrgBar.SetProg( 20 ) '�i���X�V
Call GetTrgtFileList(asAllFileList, asTrgtFileList)
'Call OutputAllElement(asTrgtFileList) '��DebugDel��

oPrgBar.SetProg( 100 ) '�i���X�V

objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime

'If 0 Then ' ��DebugDel��
' ******************************************
' * Persistent ID ���X�g�i�A�z�z��j�擾   *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** Persistent ID ���X�g�i�A�z�z��j�擾���� *** "
oPrgBar.SetMsg( _
	"�@�E���t���͏���" & vbNewLine & _
	"�@�E�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�ˁEPersistent ID ���X�g�i�A�z�z��j�擾" & vbNewLine & _ 
	"�@�E�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '�i���X�V

Dim htPath2PerID
Set htPath2PerID = CreateObject("Scripting.Dictionary")

Dim lTrackNum
lTrackNum = objTracks.Count
'lTrackNum = 1000

Dim iTrackIdx
Dim sPersistentID
Dim sLocation

On Error Resume Next
For iTrackIdx = 1 To lTrackNum
	oPrgBar.SetProg( ( iTrackIdx / lTrackNum ) * 100 ) '�i���X�V
	IF objTracks.Item( iTrackIdx ).KindAsString = "MPEG �I�[�f�B�I�t�@�C��" Then
		sPersistentID = GetPerIDFromObj( objTracks.Item( iTrackIdx ) )
		sLocation = objTracks.Item( iTrackIdx ).Location
		If htPath2PerID.Exists( sLocation ) = True Then
			'Do Nothing
		Else
			htPath2PerID.Add sLocation, sPersistentID
			objLogFile.WriteLine sLocation & Chr(9) & sPersistentID
		End If
		
		If Err.Number <> 0 Then
			MsgBox "Err.Number    :" & Err.Number & vbNewLine & _
				   "iTrackIdx     :" & iTrackIdx & vbNewLine & _
				   "Location      :" & Location & vbNewLine & _
				   "sPersistentID :" & sPersistentID
			Call ExecuteFinish
		Else
			'Do Nothing
		End If
	Else
		'Do Nothing
	End If
Next
objLogFile.WriteLine "�t�@�C�����F" & htPath2PerID.Count
Err.Clear
On Error Goto 0

objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime

' ******************************************
' * �^�O�X�V                               *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �^�O�X�V���� *** "
oPrgBar.SetMsg( _
	"�@�E���t���͏���" & vbNewLine & _
	"�@�E�X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
	"�@�EPersistent ID ���X�g�i�A�z�z��j�擾" & vbNewLine & _ 
	"�ˁE�^�O�X�V����" & vbNewLine & _
	"" _
)
oPrgBar.SetProg( 0 ) '�i���X�V

Dim lTrgtFileListIdx
Dim lTrgtFileListNum
Dim sFilePath
Dim objTrgtTrack
lTrgtFileListNum = UBound( asTrgtFileList )
For lTrgtFileListIdx = 0 To lTrgtFileListNum
	oPrgBar.SetProg( ( ( lTrgtFileListIdx + 1 ) / ( lTrgtFileListNum + 1 ) ) * 100 ) '�i���X�V
	sFilePath = asTrgtFileList( lTrgtFileListIdx )
	sPersistentID = htPath2PerID.Item( sFilePath )
	
	Set objTrgtTrack = GetObjFromPerID( sPersistentID )
	objTrgtTrack.Composer = "1"
	objTrgtTrack.Composer = ""
	Set objTrgtTrack = Nothing
Next

objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime

'End If ' ��DebugDel��

' ******************************************
' * �I������                               *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �I������ *** "

Call ExecuteFinish()

MsgBox "�v���O����������ɏI�����܂����B"

'==========================================================
'= �֐���`
'==========================================================
Function ExecuteFinish()
	
	Call oStpWtch.StopT
	
	objLogFile.WriteLine "StartPoint      : " & oStpWtch.StartPoint
	objLogFile.WriteLine "StopPoint       : " & oStpWtch.StopPoint
	objLogFile.WriteLine "ElapsedTime[s]  : " & oStpWtch.ElapsedTime
	objLogFile.WriteLine "IntervalTime[s] : " & oStpWtch.IntervalTime
	
	objLogFile.Close
	Call oPrgBar.Quit
	
	Set oStpWtch = Nothing
	Set oPrgBar = Nothing
End Function

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

Function GetTrgtFileList( _
	ByRef asAllFileList, _
	ByRef asTrgtFileList _
)
	Dim lAllFileListIdx
	Dim lTrgtFileListIdx
	Dim sExtName
	Dim sLastModDate
	Dim oFileSys
	Dim sFilePath
	
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	
	For lAllFileListIdx = 0 to UBound(asAllFileList)
		oPrgBar.SetProg( ( lAllFileListIdx / UBound(asAllFileList) ) * 100 ) '�i���X�V
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
	objLogFile.WriteLine "�t�@�C�����F" & UBound(asAllFileList) + 1
	
	Set oFileSys = Nothing
End Function
