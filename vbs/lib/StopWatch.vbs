Option Explicit

Class StopWatch
	Dim glStartPoint
	Dim glStopPoint
	Dim glIntervalTime
	
	Private Sub Class_Initialize()
		Call StopWatchInit
	End Sub
	
	'*** 初期化 ***
	Private Function StopWatchInit()
		glStartPoint = 0
		glStopPoint = 0
		glIntervalTime = 0
	End Function
	
	'*** 測定開始 ***
	Public Sub StartT()
		glStartPoint = Now()
		glIntervalTime = glStartPoint
	End Sub
	
	'*** 測定停止 ***
	Public Function StopT()
		glStopPoint = Now()
		StopT = glStopPoint - glStartPoint
	End Function
	
	'*** 開始から現在までの総経過時間[s] ***
	Public Property Get ElapsedTime()
		ElapsedTime = DateDiff( "s", glStartPoint, Now() )
	End Property
	
	'*** 前回 IntervalTime() 呼び出し時からの時間間隔[s] ***
	Public Property Get IntervalTime()
		Dim lCurTime
		lCurTime = Now()
		IntervalTime = DateDiff( "s", glIntervalTime, lCurTime )
		glIntervalTime = lCurTime
	End Property
	
	'*** 開始時刻 ***
	Public Property Get StartPoint()
		StartPoint = glStartPoint
	End Property
	
	'*** 終了時刻 ***
	Public Property Get StopPoint()
		StopPoint = glStopPoint
	End Property
	
	Private Sub Class_Terminate()
		Call StopWatchInit
	End Sub
End Class
