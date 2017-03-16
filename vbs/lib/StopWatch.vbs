Option Explicit

Private Const STPWTCH_ONEDAYSEC = 86400 '60[s] * 60[m] * 24[h]
Private Const STPWTCH_ONEHOURSEC = 3600 '60[s] * 60[m]
Private Const STPWTCH_ONEMINSEC = 60    '60[s]

'���Ԍv���������擾�i���\�~���b���x�̌덷�L��j
Class StopWatch
	Dim gbIsMeasuring
	Dim gdStartPoint
	Dim gdStopPoint
	Dim gdIntervalPoint
	Dim gsStartDate
	Dim gsStopDate
	Dim gsIntervalDate
	
	Private Sub Class_Initialize()
		Call StopWatchInit
	End Sub
	
	Private Sub Class_Terminate()
		Call StopWatchInit
	End Sub
	
	'*** ������ ***
	Private Sub StopWatchInit()
		gbIsMeasuring = False
		gdStartPoint = 0
		gdStopPoint = 0
		gdIntervalPoint = 0
		gsStartDate = Date()
		gsStopDate = Date()
		gsIntervalDate = Date()
	End Sub
	
	' ==================================================================
	' = �T�v	������J�n����
	' = ����	�Ȃ�
	' = �ߒl			String	����J�n�����i��A3:40:29 .781�j
	' = �o��	�Ȃ�
	' ==================================================================
	Public Function StartT()
		gbIsMeasuring = True
		gdStartPoint = Timer()
		gsStartDate = Date()
		gdIntervalPoint = gdStartPoint
		gsIntervalDate = gsStartDate
		StartT = ConvFormat( gdStartPoint, 1 )
	End Function
	
	' ==================================================================
	' = �T�v	������~����
	' = ����	�Ȃ�
	' = �ߒl			String	�����~�����i��A3:40:29 .781�j
	' = �o��	�Ȃ�
	' ==================================================================
	Public Function StopT()
		gbIsMeasuring = True
		gdStopPoint = Timer()
		gsStopDate = Date()
		StopT = ConvFormat( gdStopPoint, 1 )
	End Function
	
	' ==================================================================
	' = �T�v	����J�n�������擾����
	' = ����	�Ȃ�
	' = �ߒl			String	����J�n�����i��A3:40:29 .781�j
	' = �o��	�Ȃ�
	' ==================================================================
	Public Property Get StartPoint()
		StartPoint = ConvFormat( gdStartPoint, 1 )
	End Property
	
	' ==================================================================
	' = �T�v	�����~�������擾����
	' = ����	�Ȃ�
	' = �ߒl			String	�����~�����i��A3:40:29 .781�j
	' = �o��	�Ȃ�
	' ==================================================================
	Public Property Get StopPoint()
		StopPoint = ConvFormat( gdStopPoint, 1 )
	End Property
	
	' ==================================================================
	' = �T�v	�J�n����̑��o�ߎ��Ԃ��擾����
	' = ����	�Ȃ�
	' = �ߒl			String	���o�ߎ��ԁi��A3h 40m 29s 781 ms�j
	' = �o��	���肪��~����Ă���ꍇ�́u�J�n�����~�܂Łv�̑��o��
	' = 		���ԁA���肪��~����Ă��Ȃ��ꍇ�́u�J�n���猻�݂܂Łv��
	' = 		���o�ߎ��Ԃ�ԋp����
	' ==================================================================
	Public Property Get ElapsedTime()
		If gbIsMeasuring = True Then
			ElapsedTime = ConvFormat( TimeDiff( gsStartDate, gdStartPoint, Date(), Timer() ), 2 )
		Else
			ElapsedTime = ConvFormat( TimeDiff( gsStartDate, gdStartPoint, gsStopDate, gdStopPoint ), 2 )
		End If
	End Property
	
	' ==================================================================
	' = �T�v	�O�� IntervalTime() �Ăяo��������̎��ԊԊu���擾����
	' = ����	�Ȃ�
	' = �ߒl			String	���o�ߎ��ԁi��A3h 40m 29s 781 ms�j
	' = �o��	�Ȃ�
	' ==================================================================
	Public Property Get IntervalTime()
		Dim dCurPoint
		Dim sCurDate
		dCurPoint = Timer()
		sCurDate = Date()
		IntervalTime = ConvFormat( TimeDiff( gsIntervalDate, gdIntervalPoint, sCurDate, dCurPoint ), 2 )
		gdIntervalPoint = dCurPoint
		gsIntervalDate = sCurDate
	End Property
	
	' ==================================================================
	' = �T�v	�������r���āA����b�ɕϊ����ĕԋp����
	' = ����	sPreDate	String	[in]	�O�̓��t
	' = ����	dPreTime	Double	[in]	�O�̎���
	' = ����	sPostDate	String	[in]	��̓��t
	' = ����	dPostTime	Double	[in]	��̎���
	' = �ߒl				Double			�����̍��فi�b�j
	' = �o��	�Ȃ�
	' ==================================================================
	Public Function TimeDiff( _
		ByVal sPreDate, _
		ByVal dPreTime, _
		ByVal sPostDate, _
		ByVal dPostTime _
	)
		Dim lDateDiff
		lDateDiff = DateDiff("d", sPreDate, sPostDate)
		If lDateDiff > 0 Then
			TimeDiff = (STPWTCH_ONEDAYSEC * (lDateDiff - 1)) + (STPWTCH_ONEDAYSEC - dPreTime) + dPostTime
		ElseIf lDateDiff = 0 Then
			TimeDiff = dPostTime - dPreTime
		Else
			TimeDiff = 0
		End If
	End Function
	
	' ==================================================================
	' = �T�v	Timer()�֐��̕ԋp�l�������`���ɕϊ�
	' =			  ex) 49229.781 �� 13:40:29 .781
	' = ����	dTimeValue	Double	[in]	�o�ߕb��
	' = ����	lTimeFormat	Long	[in]	�����`��
	' = 										1) 3:40:29 .781
	' = 										2) 3h 40m 29s 781 ms
	' = �ߒl				String			�ϊ�����
	' = �o��	�Ȃ�
	' ==================================================================
	Public Function ConvFormat( _
		ByVal dTimeValue, _
		ByVal lTimeFormat _
	)
		Dim lTime
		lTime = Fix( dTimeValue )
		
		Dim lTemp
		Dim lHour
		Dim lMinite
		Dim lSecond
		Dim lMinSec
		
		'�����Z�o
		lHour = Fix( lTime / STPWTCH_ONEHOURSEC )
		lTemp = Fix( lTime Mod STPWTCH_ONEHOURSEC )
		lMinite = Fix( lTemp / STPWTCH_ONEMINSEC )
		lSecond = Fix( lTemp Mod STPWTCH_ONEMINSEC )
		lMinSec = Round( dTimeValue - lTime, 3 )
		lMinSec = Mid( CStr(lMinSec), 3, 3 )
		lMinSec = lMinSec & String( 3 - Len(lMinSec), "0" )
		
		If lTimeFormat = 1 Then
			lHour = String( 2 - Len(CStr(lHour)),"0" ) & lHour
			lMinite = String( 2 - Len(CStr(lMinite)), "0" ) & lMinite
			lSecond = String( 2 - Len(CStr(lSecond)), "0" ) & lSecond
			ConvFormat = lHour & ":" & lMinite & ":" & lSecond & " ." & lMinSec
		ElseIf lTimeFormat = 2 Then
			If lHour = 0 Then
				If lMinite = 0 Then
					ConvFormat = lSecond & "[s] " & lMinSec & "[ms]"
				Else
					If lSecond = 0 Then
						ConvFormat = lMinSec & "[ms]"
					Else
						ConvFormat = lSecond & "[s] " & lMinSec & "[ms]"
					End If
				End If
			Else
				ConvFormat = lHour & "[h] " & lMinite & "[m] " & lSecond & "[s] " & lMinSec & "[ms]"
			End If
		Else
			ConvFormat = ""
		End If
	End Function
End Class
	If WScript.ScriptName = "StopWatch.vbs" Then
		Call Test_StopWatch
	End If
	Private Sub Test_StopWatch
		Dim oStpWtch
		Set oStpWtch = New StopWatch
		Dim bTestContinue
		Dim bAllTestExec
		Dim bIsTestFinish
		
		Dim lTestCase
		lTestCase = InputBox("�e�X�g�P�[�X����͂��Ă�������")
		
		If lTestCase = 0 Then
			bAllTestExec = True
		Else
			bAllTestExec = False
		End If
		
		bIsTestFinish = False
		bTestContinue = True
		Do While bTestContinue = True
			Dim sOutMsg
			sOutMsg = ""
			Select Case lTestCase
				Case 0
					'Do Nothing
				Case 1:
					sOutMsg = sOutMsg & vbNewLine & "### start! ###"
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.StartT
					
					WScript.Sleep 1000
					sOutMsg = sOutMsg & vbNewLine & "--- wait1000 ---"
					sOutMsg = sOutMsg & vbNewLine & "StartPoint   : " & oStpWtch.StartPoint
					sOutMsg = sOutMsg & vbNewLine & "StopPoint    : " & oStpWtch.StopPoint
					sOutMsg = sOutMsg & vbNewLine & "IntervalTime : " & oStpWtch.IntervalTime
					sOutMsg = sOutMsg & vbNewLine & "ElapsedTime  : " & oStpWtch.ElapsedTime
					
					WScript.Sleep 1000
					sOutMsg = sOutMsg & vbNewLine & "--- wait1000 ---"
					sOutMsg = sOutMsg & vbNewLine & "StartPoint   : " & oStpWtch.StartPoint
					sOutMsg = sOutMsg & vbNewLine & "StopPoint    : " & oStpWtch.StopPoint
					sOutMsg = sOutMsg & vbNewLine & "IntervalTime : " & oStpWtch.IntervalTime
					sOutMsg = sOutMsg & vbNewLine & "ElapsedTime  : " & oStpWtch.ElapsedTime
					
					sOutMsg = sOutMsg & vbNewLine & "### stop! ###"
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.StopT
					
					sOutMsg = sOutMsg & vbNewLine & "StartPoint   : " & oStpWtch.StartPoint
					sOutMsg = sOutMsg & vbNewLine & "StopPoint    : " & oStpWtch.StopPoint
					sOutMsg = sOutMsg & vbNewLine & "IntervalTime : " & oStpWtch.IntervalTime
					sOutMsg = sOutMsg & vbNewLine & "ElapsedTime  : " & oStpWtch.ElapsedTime
				Case 2:
					sOutMsg = sOutMsg & vbNewLine & "### start! ###"
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.StartT
					
					WScript.Sleep 2000
					sOutMsg = sOutMsg & vbNewLine & "--- wait2000 ---"
					sOutMsg = sOutMsg & vbNewLine & "StartPoint   : " & oStpWtch.StartPoint
					sOutMsg = sOutMsg & vbNewLine & "StopPoint    : " & oStpWtch.StopPoint
					sOutMsg = sOutMsg & vbNewLine & "IntervalTime : " & oStpWtch.IntervalTime
					sOutMsg = sOutMsg & vbNewLine & "ElapsedTime  : " & oStpWtch.ElapsedTime
					
					WScript.Sleep 4000
					sOutMsg = sOutMsg & vbNewLine & "--- wait4000 ---"
					sOutMsg = sOutMsg & vbNewLine & "StartPoint   : " & oStpWtch.StartPoint
					sOutMsg = sOutMsg & vbNewLine & "StopPoint    : " & oStpWtch.StopPoint
					sOutMsg = sOutMsg & vbNewLine & "IntervalTime : " & oStpWtch.IntervalTime
					sOutMsg = sOutMsg & vbNewLine & "ElapsedTime  : " & oStpWtch.ElapsedTime
					
					sOutMsg = sOutMsg & vbNewLine & "### stop! ###"
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.StopT
					
					sOutMsg = sOutMsg & vbNewLine & "StartPoint   : " & oStpWtch.StartPoint
					sOutMsg = sOutMsg & vbNewLine & "StopPoint    : " & oStpWtch.StopPoint
					sOutMsg = sOutMsg & vbNewLine & "IntervalTime : " & oStpWtch.IntervalTime
					sOutMsg = sOutMsg & vbNewLine & "ElapsedTime  : " & oStpWtch.ElapsedTime
				Case 3:
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/13", 6003 ), 1 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/13", 6003 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & ""
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 0,    "2016/12/11", 6003 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/11", 6003 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/11", 6059 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/11", 6060 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/11", 9599 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/11", 9600 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/12", 5999 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/12", 6000 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & ""
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 0.12, "2016/12/11", 0.59 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & ""
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/12", 6003 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/13", 6003 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & ""
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/11", 6000 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/11", 5999 ), 2 )
					sOutMsg = sOutMsg & vbNewLine & oStpWtch.ConvFormat( oStpWtch.TimeDiff( "2016/12/11", 6000, "2016/12/10", 6003 ), 2 )
				Case Else:
					bIsTestFinish = True
			End Select
			MsgBox sOutMsg
		
			If bAllTestExec = True Then
				If bIsTestFinish = True Then
					bTestContinue = False
				Else
					lTestCase = lTestCase + 1
					bTestContinue = True
				End If
			Else
				bTestContinue = False
			End If
		Loop
		
		Set oStpWtch = Nothing
	End Sub
