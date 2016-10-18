Option Explicit

Const ONEDAYSEC = 86400 '60[s] * 60[m] * 24[h]
Const ONEHOURSEC = 3600 '60[s] * 60[m]
Const ONEMINSEC = 60    '60[s]

'���Ԍv���������擾�i���\�~���b���x�̌덷�L��j
Class StopWatch
	Dim glStartPoint
	Dim glStopPoint
	Dim glIntervalTime
	
	Private Sub Class_Initialize()
		Call StopWatchInit
	End Sub
	
	'*** ������ ***
	Private Function StopWatchInit()
		glStartPoint = 0
		glStopPoint = 0
		glIntervalTime = 0
	End Function
	
	'*** ����J�n ***
	Public Sub StartT()
		glStartPoint = Timer()
		glIntervalTime = glStartPoint
	End Sub
	
	'*** �����~ ***
	Public Function StopT()
		glStopPoint = Timer()
		StopT = glStopPoint - glStartPoint
	End Function
	
	'*** �J�n���猻�݂܂ł̑��o�ߎ���[s] ***
	Public Property Get ElapsedTime()
		ElapsedTime = TimeDiff( glStartPoint, Timer() )
	End Property
	
	'*** �O�� IntervalTime() �Ăяo��������̎��ԊԊu[s] ***
	Public Property Get IntervalTime()
		Dim lCurTime
		lCurTime = Timer()
		IntervalTime = TimeDiff( glIntervalTime, lCurTime )
		glIntervalTime = lCurTime
	End Property
	
	'*** �J�n���� ***
	Public Property Get StartPoint()
		StartPoint = ConvFormat( glStartPoint )
	End Property
	
	'*** �I������ ***
	Public Property Get StopPoint()
		StopPoint = ConvFormat( glStopPoint )
	End Property
	
	Private Sub Class_Terminate()
		Call StopWatchInit
	End Sub
	
	Private Function TimeDiff( _
		ByVal lPreTime, _
		ByVal lPostTime _
	)
		If lPostTime >= lPreTime Then
			TimeDiff = lPostTime - lPreTime
		Else
			' �^�钆��0�����ׂ����Ƃ��̑Ώ�
			TimeDiff = ( ONEDAYSEC - lPreTime ) + lPostTime
		End If
	End Function
	
	' Timer()�֐��̕ԋp�l�������`���ɕϊ�
	'   ex) 49229.781 �� 13:40:29 .781
	Private Function ConvFormat( _
		ByVal lTimeValue _
	)
		Dim lHour
		Dim lMinite
		Dim lSecond
		Dim lTime
		Dim lMinSec
		
		lTime = Fix( lTimeValue )
		
		'�����Z�o
		lHour = Fix( lTime / ONEHOURSEC )
		lMinite = Fix( ( lTime - ( lHour * ONEHOURSEC ) ) / ONEMINSEC )
		lSecond = Fix( ( lTime - ( lHour * ONEHOURSEC ) - ( lMinite * ONEMINSEC ) ) )
		lMinSec = Round( lTimeValue - lTime, 3 )
		
		'�����񐮌`
		lHour = String( 2 - Len(CStr(lHour)),   "0" ) & lHour
		lMinite = String( 2 - Len(CStr(lMinite)), "0" ) & lMinite
		lSecond = String( 2 - Len(CStr(lSecond)), "0" ) & lSecond
		lMinSec = Mid( CStr(lMinSec), 3, 3 )
		lMinSec = lMinSec & String( 3 - Len(lMinSec), "0" )
		
		ConvFormat = lHour & ":" & lMinite & ":" & lSecond & " ." & lMinSec
	End Function
End Class

'Call Test
'Private Sub Test
'	Dim oStpWtch
'	Set oStpWtch = New StopWatch
'	
'	Dim sTime
'	
'	sTime = ""
'	oStpWtch.StartT
'	
'	WScript.Sleep 1000
'	sTime = sTime & vbNewLine & oStpWtch.IntervalTime
'	sTime = sTime & vbNewLine & oStpWtch.ElapsedTime
'	sTime = sTime & vbNewLine & ""
'	WScript.Sleep 1000
''	sTime = sTime & vbNewLine & oStpWtch.IntervalTime
''	WScript.Sleep 100000
''	sTime = sTime & vbNewLine & oStpWtch.ElapsedTime
''	sTime = sTime & vbNewLine & oStpWtch.IntervalTime
''	WScript.Sleep 2000
''	sTime = sTime & vbNewLine & oStpWtch.ElapsedTime
''	sTime = sTime & vbNewLine & oStpWtch.IntervalTime
''	WScript.Sleep 4000
''	sTime = sTime & vbNewLine & oStpWtch.StartPoint
''	sTime = sTime & vbNewLine & oStpWtch.StopPoint
''	sTime = sTime & vbNewLine & oStpWtch.ElapsedTime
''	sTime = sTime & vbNewLine & oStpWtch.IntervalTime
'	
'	oStpWtch.StopT
'	
'	sTime = sTime & vbNewLine & oStpWtch.IntervalTime
'	sTime = sTime & vbNewLine & oStpWtch.ElapsedTime
'	sTime = sTime & vbNewLine & oStpWtch.StartPoint
'	sTime = sTime & vbNewLine & oStpWtch.StopPoint
'	
'	MsgBox sTime
'End Sub
