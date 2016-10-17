Option Explicit

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
		glStartPoint = Now()
		glIntervalTime = glStartPoint
	End Sub
	
	'*** �����~ ***
	Public Function StopT()
		glStopPoint = Now()
		StopT = glStopPoint - glStartPoint
	End Function
	
	'*** �J�n���猻�݂܂ł̑��o�ߎ���[s] ***
	Public Property Get ElapsedTime()
		ElapsedTime = DateDiff( "s", glStartPoint, Now() )
	End Property
	
	'*** �O�� IntervalTime() �Ăяo��������̎��ԊԊu[s] ***
	Public Property Get IntervalTime()
		Dim lCurTime
		lCurTime = Now()
		IntervalTime = DateDiff( "s", glIntervalTime, lCurTime )
		glIntervalTime = lCurTime
	End Property
	
	'*** �J�n���� ***
	Public Property Get StartPoint()
		StartPoint = glStartPoint
	End Property
	
	'*** �I������ ***
	Public Property Get StopPoint()
		StopPoint = glStopPoint
	End Property
	
	Private Sub Class_Terminate()
		Call StopWatchInit
	End Sub
End Class
