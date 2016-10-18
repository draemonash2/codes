Option Explicit

Const ONEDAYSEC = 86400 '60[s] * 60[m] * 24[h]
Const ONEHOURSEC = 3600 '60[s] * 60[m]
Const ONEMINSEC = 60    '60[s]

'時間計測＆時刻取得（数十ミリ秒程度の誤差有り）
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
		glStartPoint = Timer()
		glIntervalTime = glStartPoint
	End Sub
	
	'*** 測定停止 ***
	Public Function StopT()
		glStopPoint = Timer()
		StopT = glStopPoint - glStartPoint
	End Function
	
	'*** 開始から現在までの総経過時間[s] ***
	Public Property Get ElapsedTime()
		ElapsedTime = TimeDiff( glStartPoint, Timer() )
	End Property
	
	'*** 前回 IntervalTime() 呼び出し時からの時間間隔[s] ***
	Public Property Get IntervalTime()
		Dim lCurTime
		lCurTime = Timer()
		IntervalTime = TimeDiff( glIntervalTime, lCurTime )
		glIntervalTime = lCurTime
	End Property
	
	'*** 開始時刻 ***
	Public Property Get StartPoint()
		StartPoint = ConvFormat( glStartPoint )
	End Property
	
	'*** 終了時刻 ***
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
			' 真夜中の0時を跨いだときの対処
			TimeDiff = ( ONEDAYSEC - lPreTime ) + lPostTime
		End If
	End Function
	
	' Timer()関数の返却値を時刻形式に変換
	'   ex) 49229.781 ⇒ 13:40:29 .781
	Private Function ConvFormat( _
		ByVal lTimeValue _
	)
		Dim lHour
		Dim lMinite
		Dim lSecond
		Dim lTime
		Dim lMinSec
		
		lTime = Fix( lTimeValue )
		
		'時刻算出
		lHour = Fix( lTime / ONEHOURSEC )
		lMinite = Fix( ( lTime - ( lHour * ONEHOURSEC ) ) / ONEMINSEC )
		lSecond = Fix( ( lTime - ( lHour * ONEHOURSEC ) - ( lMinite * ONEMINSEC ) ) )
		lMinSec = Round( lTimeValue - lTime, 3 )
		
		'文字列整形
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
