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

If 1 Then ' ��Debug��

' ******************************************
' * ���t����                               *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** ���t���͏��� *** "
oPrgBar.SetMsg( _
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

' ******************************************
' * �X�V�Ώۃt�@�C�����X�g�擾             *
' ******************************************
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �X�V�Ώۃt�@�C������ *** "
oPrgBar.SetMsg( _
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
sExecCmd = "Dir """ & TRGT_DIR & """ /s /a:a-d > """ & sTmpFilePath & """"
With CreateObject("Wscript.Shell")	
	.Run "cmd /c" & sExecCmd, 7, True
End With
Dim objFile
Dim sTextAll
Dim asTxtArray
If Err.Number = 0 Then
	Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
	If Err.Number = 0 Then
		sTextAll = objFile.ReadAll
		sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '�����ɉ��s���t�^����Ă��܂����߁A�폜
		asTxtArray = Split( sTextAll, vbNewLine )
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

'*** Dir �R�}���h���ʎ擾 �� �X�V�Ώۃt�@�C�����X�g�쐬 ***
Dim lIdx
Dim sDirPath
Dim sTxtLine
Dim sModDate
Dim sFileName
Dim sFilePath
Dim sFileSize
Dim vSplitData
Dim asTrgtFileList()
ReDim asTrgtFileList(-1)
sDirPath = ""
For lIdx = 0 to UBound( asTxtArray )
	oPrgBar.SetProg( oPrgBar.ConvProgRange( 0, UBound( asTxtArray ), lIdx ) ) '�i���X�V
	sTxtLine = asTxtArray( lIdx )
	If InStr( sTxtLine, " �̃f�B���N�g��" ) > 0 Then
		sDirPath = Mid( sTxtLine, 2, Len( sTxtLine ) - Len( " �̃f�B���N�g��" ) - 1 )
	ElseIf InStr( sTxtLine, "�{�����[�� ���x��" ) > 0 Or _
		   InStr( sTxtLine, "�{�����[�� �V���A���ԍ��� " ) > 0 Or _
		   ( ( InStr( sTxtLine, " �̃t�@�C��" ) > 0 ) And ( InStr( sTxtLine, " �o�C�g" ) > 0 ) ) Or _
		   ( ( InStr( sTxtLine, " �̃f�B���N�g��" ) > 0 ) And ( InStr( sTxtLine, " �o�C�g" ) > 0 ) ) Or _
		   InStr( sTxtLine, "     �t�@�C���̑���:" ) > 0 Or _
		   sTxtLine = "" Then
		'Do Nothing
	Else
		Do
			sTxtLine = Replace( sTxtLine, "  ", " " )
		Loop While InStr( sTxtLine, "  " ) > 0
		vSplitData = Split( sTxtLine, " " )
		If UBound( vSplitData ) < 3 Then
			MsgBox "�G���[�I" & vbNewLine & sTxtLine
			WScript.Quit
		End If
		sModDate = vSplitData(0) & " " & vSplitData(1) & ":00"
		sFileSize = vSplitData(2)
		Dim i
		For i = 3 to UBound( vSplitData )
			If i = 3 Then
				sFileName = vSplitData(i)
			Else
				sFileName = sFileName & " " & vSplitData(i)
			End If
		Next
		sFilePath = sDirPath & "\" & sFileName
		
		If DateDiff("s", sCmpBaseTime, sModDate ) >= 0  Then
			ReDim Preserve asTrgtFileList( UBound( asTrgtFileList ) + 1 )
			asTrgtFileList( UBound( asTrgtFileList ) ) = sFilePath
		Else
			'Do Nothing
		End If
		
		If Err.Number = 0 Then
			'Do Nothing
		Else
			MsgBox "sFilePath  : " & sFilePath & vbNewLine & _
				   "sModDate   : " & sModDate  & vbNewLine & _
				   "sFileSize  : " & sFileSize & vbNewLine & _
				   "sFileName  : " & sFileName & vbNewLine & _
				   "ArrayData  : " & asTrgtFileList( UBound( asTrgtFileList ) )
			MsgBox Err.Description
			Err.Clear
		End If
	End If
Next

objFSO.DeleteFile sTmpFilePath, True
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
objLogFile.WriteLine ""
objLogFile.WriteLine "*** �^�O�X�V���� *** "
oPrgBar.SetMsg( _
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
