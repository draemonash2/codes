Option Explicit

Const TRGT_DIR = "Z:\300_Musics"
 
Dim objItunes
Dim objTracks
Dim oPrgBar
Dim oStpWtch

Set objItunes = WScript.CreateObject("iTunes.Application")
Set objTracks = objItunes.LibraryPlaylist.Tracks
Set oPrgBar = New ProgressBar
Set oStpWtch = New StopWatch

Dim objWshShell
Dim objLogFile
Dim sLogFilePath
Dim objFso

Set objFso = CreateObject("Scripting.FileSystemObject")
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
sLogFilePath = objWshShell.CurrentDirectory & "\" & _
               RemoveTailWord( WScript.ScriptName, "." ) & "_log.txt"

Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )

Call oStpWtch.StartT

objLogFile.WriteLine "[���ݎ���] " & Now()

' ******************************************
' * ���t����                               *
' ******************************************
objLogFile.WriteLine ""
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

objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime

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

objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime

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
		sPersistentID = PersistentID( objTracks.Item( iTrackIdx ) )
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
objLogFile.WriteLine "�擾�����t�@�C�����F" & htPath2PerID.Count
Err.Clear
On Error Goto 0

objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime

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
lTrgtFileListNum = UBound( asTrgtFileList )
For lTrgtFileListIdx = 0 To lTrgtFileListNum
	oPrgBar.SetProg( ( ( lTrgtFileListIdx + 1 ) / ( lTrgtFileListNum + 1 ) ) * 100 ) '�i���X�V
	sFilePath = asTrgtFileList( lTrgtFileListIdx )
	sPersistentID = htPath2PerID.Item( sFilePath )
	
	ObjectFromID( sPersistentID ).Composer = "1"
	ObjectFromID( sPersistentID ).Composer = ""
Next

objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime

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

	objLogFile.WriteLine "StartTime    : " & oStpWtch.StartTime
	objLogFile.WriteLine "StopTime     : " & oStpWtch.StopTime
	objLogFile.WriteLine "LapTime      : " & oStpWtch.LapTime
	objLogFile.WriteLine "IntervalTime : " & oStpWtch.IntervalTime
	
	objLogFile.Close
	Call oPrgBar.Quit

	Set oStpWtch = Nothing
	Set oPrgBar = Nothing

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
	
	Set oFileSys = Nothing
End Function

Function OutputAllElement( _
	ByRef asOutTrgtArray _
)
	Dim lIdx
	Dim sOutStr
	sOutStr = "EleNum :" & Ubound( asOutTrgtArray ) + 1
	For lIdx = 0 to UBound( asOutTrgtArray )
		sOutStr = sOutStr & vbNewLine & asOutTrgtArray(lIdx)
	Next
	WScript.Echo sOutStr
End Function

' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕������ԋp����B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ���o������
' = �o��    �Ȃ�
' ==================================================================
Public Function ExtractTailWord( _
    ByVal sStr, _
    ByVal sDlmtr _
)
    Dim asSplitWord
 
    If Len(sStr) = 0 Then
        ExtractTailWord = ""
    Else
        ExtractTailWord = ""
        asSplitWord = Split(sStr, sDlmtr)
        ExtractTailWord = asSplitWord(UBound(asSplitWord))
    End If
End Function
 
' ==================================================================
' = �T�v    ������؂蕶���ȍ~�̕��������������B
' = ����    sStr        String  [in]  �������镶����
' = ����    sDlmtr      String  [in]  ��؂蕶��
' = �ߒl                String        ����������
' = �o��    �Ȃ�
' ==================================================================
Public Function RemoveTailWord( _
    ByVal sStr, _
    ByVal sDlmtr _
)
    Dim sTailWord
    Dim lRemoveLen
 
    If sStr = "" Then
        RemoveTailWord = ""
    Else
        If sDlmtr = "" Then
            RemoveTailWord = sStr
        Else
            If InStr(sStr, sDlmtr) = 0 Then
                RemoveTailWord = sStr
            Else
                sTailWord = ExtractTailWord(sStr, sDlmtr)
                lRemoveLen = Len(sDlmtr) + Len(sTailWord)
                RemoveTailWord = Left(sStr, Len(sStr) - lRemoveLen)
            End If
        End If
    End If
End Function

'lFileListType�j0�F�����A1:�t�@�C���A2:�t�H���_�A����ȊO�F�i�[���Ȃ�
Function GetFileList( _
	ByVal sTrgtDir, _
	ByRef asFileList, _
	ByVal lFileListType _
)
	Dim objFileSys
	Dim objFolder
	Dim objSubFolder
	Dim objFile
	Dim bExecStore
	Dim lLastIdx
 
	Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFileSys.GetFolder( sTrgtDir )
 
	'*** �t�H���_�p�X�i�[ ***
	Select Case lFileListType
		Case 0:    bExecStore = True
		Case 1:    bExecStore = False
		Case 2:    bExecStore = True
		Case Else: bExecStore = False
	End Select
	If bExecStore = True Then
		lLastIdx = UBound( asFileList ) + 1
		ReDim Preserve asFileList( lLastIdx )
		asFileList( lLastIdx ) = objFolder
	Else
		'Do Nothing
	End If
 
	'�t�H���_���̃T�u�t�H���_���
	'�i�T�u�t�H���_���Ȃ���΃��[�v���͒ʂ�Ȃ��j
	For Each objSubFolder In objFolder.SubFolders
		Call GetFileList( objSubFolder, asFileList, lFileListType)
	Next
 
	'*** �t�@�C���p�X�i�[ ***
	For Each objFile In objFolder.Files
		Select Case lFileListType
			Case 0:    bExecStore = True
			Case 1:    bExecStore = True
			Case 2:    bExecStore = False
			Case Else: bExecStore = False
		End Select
		If bExecStore = True Then
			'�{�X�N���v�g�t�@�C���͊i�[�ΏۊO
			If objFile.Name = WScript.ScriptName Then
				'Do Nothing
			Else
				lLastIdx = UBound( asFileList ) + 1
				ReDim Preserve asFileList( lLastIdx )
				asFileList( lLastIdx ) = objFile
			End If
		Else
			'Do Nothing
		End If
	Next
 
	Set objFolder = Nothing
	Set objFileSys = Nothing
End Function

Class ProgressBar
	Dim objExplorer
	Dim sProgressMsg
	
	Private Sub Class_Initialize()
		Dim objWMIService
		Dim colItems
		Dim strComputer
		Dim objItem
		Dim intHorizontal
		Dim intVertical
		strComputer = "."
		Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
		For Each objItem in colItems
			intHorizontal = objItem.ScreenWidth
			intVertical = objItem.ScreenHeight
		Next
		Set objWMIService = Nothing
		Set colItems = Nothing
		
		Set objExplorer = CreateObject("InternetExplorer.Application")
		objExplorer.Navigate "about:blank"
		objExplorer.ToolBar = 0
		objExplorer.StatusBar = 0
		objExplorer.Left = (intHorizontal - 400) / 2
		objExplorer.Top = (intVertical - 200) / 2
		objExplorer.Width = 400
		objExplorer.Height = 500
		objExplorer.Visible = 1
		
		Call ActiveIE
		objExplorer.Document.Body.Style.Cursor = "wait"
		objExplorer.Document.Title = "�i����"
		SetProg(0)
	End Sub
	
	Private Sub Class_Terminate()
		'Do Nothing
	End Sub
	
	Public Function SetMsg( _
		ByVal sMessage _
	)
		'���s������<br>�ɒu��
		sProgressMsg = Replace( sMessage, vbNewLine, "<br>" )
	End Function
	
	Public Function SetProg( _
		ByVal lProgress _
	)
		Dim lProgress100
		Dim lProgress10
	
		If lProgress > 100 Or lProgress < 0 Then
			MsgBox "�v���O���X�o�[�̐i���ɋK��l[0-100]�O�̒l���w�肳��Ă��܂��I" & vbNewLine & _
				   "�l�F" & lProgress
			MsgBox "�v���O�����𒆎~���܂��I"
			Call ProgressBarQuit
			WScript.Quit
		End If
		
		lProgress100 = Fix(lProgress)
		lProgress10 = Fix(lProgress / 10)
		
		objExplorer.Document.Body.InnerHTML = sProgressMsg & "<br>" & _
											  "<br>" & _
											  "������..." & "<br>" & _
											  String( lProgress10, "��") & String( 10 - lProgress10, "��") & _
											  "  " & lProgress100 & "% ����"
	End Function
	
	Public Function Quit()
		objExplorer.Document.Body.Style.Cursor = "default"
		objExplorer.Quit
	End Function
	
	Private Function ActiveIE()
		Dim objWshShell
		Dim intProcID
	
		Const strIEexe = "iexplore.exe" 'IE�̃v���Z�X��
		intProcID = GetProcID(strIEexe)
		Set objWshShell = CreateObject("Wscript.Shell")
		objWshShell.AppActivate intProcID
		Set objWshShell = Nothing
	End Function
	
	Private Function GetProcID(ProcessName)
		Dim Service
		Dim QfeSet
		Dim Qfe
		Dim intProcID
		
		Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
		Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='"& ProcessName &"'")
		
		intProcID = 0
		
		For Each Qfe in QfeSet
			intProcID = Qfe.ProcessId
			GetProcID = intProcID
			Exit For
		Next
	End Function
End Class

Function ObjectFromID( sID )
	Set ObjectFromID = objItunes.LibraryPlaylist.Tracks.ItemByPersistentID(Eval("&H" & Left( sID, 8)),Eval("&H" & Right( sID, 8)))
End Function

Function PersistentID( objT )
	PersistentID = Right("0000000" & Hex(objItunes.ITObjectPersistentIDHigh(objT)),8) & _
	               Right("0000000" & Hex(objItunes.ITObjectPersistentIDLow(objT)),8)
End Function

Class StopWatch
	Dim glStartTime
	Dim glStopTime
	Dim glIntervalTime
	
	Private Sub Class_Initialize()
		Call StopWatchInit
	End Sub
	
	'*** ������ ***
	Private Function StopWatchInit()
		glStartTime = 0
		glStopTime = 0
		glIntervalTime = 0
	End Function
	
	'*** ����J�n ***
	Public Sub StartT()
		glStartTime = Now()
		glIntervalTime = glStartTime
	End Sub
	
	'*** �����~ ***
	Public Function StopT()
		glStopTime = Now()
		StopT = glStopTime - glStartTime
	End Function
	
	'*** �J�n���猻�݂܂ł̑��o�ߎ���[s] ***
	Public Property Get LapTime()
'		LapTime = Now() - glStartTime
		LapTime = DateDiff( "s", glStartTime, Now() )
	End Property
	
	'*** �O�� IntervalTime() �Ăяo��������̎��ԊԊu[s] ***
	Public Property Get IntervalTime()
		Dim lCurTime
		lCurTime = Now()
'		IntervalTime = lCurTime - glIntervalTime
		IntervalTime = DateDiff( "s", glIntervalTime, lCurTime )
		glIntervalTime = lCurTime
	End Property
	
	'*** �J�n���� ***
	Public Property Get StartTime()
		StartTime = glStartTime
	End Property
	
	'*** �I������ ***
	Public Property Get StopTime()
		StopTime = glStopTime
	End Property
	
	Private Sub Class_Terminate()
		Call StopWatchInit
	End Sub
End Class
