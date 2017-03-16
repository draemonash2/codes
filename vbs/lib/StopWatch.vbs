Option Explicit

Private Const STPWTCH_ONEDAYSEC = 86400 '60[s] * 60[m] * 24[h]
Private Const STPWTCH_ONEHOURSEC = 3600 '60[s] * 60[m]
Private Const STPWTCH_ONEMINSEC = 60    '60[s]

'時間計測＆時刻取得（数十ミリ秒程度の誤差有り）
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
	
	'*** 初期化 ***
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
	' = 概要	測定を開始する
	' = 引数	なし
	' = 戻値			String	測定開始時刻（例、3:40:29 .781）
	' = 覚書	なし
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
	' = 概要	測定を停止する
	' = 引数	なし
	' = 戻値			String	測定停止時刻（例、3:40:29 .781）
	' = 覚書	なし
	' ==================================================================
	Public Function StopT()
		gbIsMeasuring = True
		gdStopPoint = Timer()
		gsStopDate = Date()
		StopT = ConvFormat( gdStopPoint, 1 )
	End Function
	
	' ==================================================================
	' = 概要	測定開始時刻を取得する
	' = 引数	なし
	' = 戻値			String	測定開始時刻（例、3:40:29 .781）
	' = 覚書	なし
	' ==================================================================
	Public Property Get StartPoint()
		StartPoint = ConvFormat( gdStartPoint, 1 )
	End Property
	
	' ==================================================================
	' = 概要	測定停止時刻を取得する
	' = 引数	なし
	' = 戻値			String	測定停止時刻（例、3:40:29 .781）
	' = 覚書	なし
	' ==================================================================
	Public Property Get StopPoint()
		StopPoint = ConvFormat( gdStopPoint, 1 )
	End Property
	
	' ==================================================================
	' = 概要	開始からの総経過時間を取得する
	' = 引数	なし
	' = 戻値			String	総経過時間（例、3h 40m 29s 781 ms）
	' = 覚書	測定が停止されている場合は「開始から停止まで」の総経過
	' = 		時間、測定が停止されていない場合は「開始から現在まで」の
	' = 		総経過時間を返却する
	' ==================================================================
	Public Property Get ElapsedTime()
		If gbIsMeasuring = True Then
			ElapsedTime = ConvFormat( TimeDiff( gsStartDate, gdStartPoint, Date(), Timer() ), 2 )
		Else
			ElapsedTime = ConvFormat( TimeDiff( gsStartDate, gdStartPoint, gsStopDate, gdStopPoint ), 2 )
		End If
	End Property
	
	' ==================================================================
	' = 概要	前回 IntervalTime() 呼び出し時からの時間間隔を取得する
	' = 引数	なし
	' = 戻値			String	総経過時間（例、3h 40m 29s 781 ms）
	' = 覚書	なし
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
	' = 概要	日時を比較して、差を秒に変換して返却する
	' = 引数	sPreDate	String	[in]	前の日付
	' = 引数	dPreTime	Double	[in]	前の時刻
	' = 引数	sPostDate	String	[in]	後の日付
	' = 引数	dPostTime	Double	[in]	後の時刻
	' = 戻値				Double			日時の差異（秒）
	' = 覚書	なし
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
	' = 概要	Timer()関数の返却値を時刻形式に変換
	' =			  ex) 49229.781 ⇒ 13:40:29 .781
	' = 引数	dTimeValue	Double	[in]	経過秒数
	' = 引数	lTimeFormat	Long	[in]	時刻形式
	' = 										1) 3:40:29 .781
	' = 										2) 3h 40m 29s 781 ms
	' = 戻値				String			変換結果
	' = 覚書	なし
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
		
		'時刻算出
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
		lTestCase = InputBox("テストケースを入力してください")
		
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
