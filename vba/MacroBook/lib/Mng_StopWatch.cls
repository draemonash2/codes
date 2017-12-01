VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mng_StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Stop Watch class v1.0
'
'<<usage sample>>
'    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'    Sub test111()
'        Dim clStopWatch As New StopWatch
'        Call clStopWatch.StartT
'
'        Sleep 1000 'ms 単位
'        Debug.Print "Lap time is " & clStopWatch.LapTime
'        Debug.Print "Interval time is " & clStopWatch.IntervalTime
'
'        Sleep 1500 'ms 単位
'        Debug.Print "Lap time is " & clStopWatch.LapTime
'        Debug.Print "Interval time is " & clStopWatch.IntervalTime
'
'        Sleep 2000 'ms 単位
'        Debug.Print "Lap time is " & clStopWatch.LapTime
'        Debug.Print "Interval time is " & clStopWatch.IntervalTime
'
'        Sleep 2500 'ms 単位
'        Debug.Print clStopWatch.StopT
'    End Sub

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
 
Dim glStartTime As Long
Dim glIntervalTime As Long
Dim gbIsTimerStarted As Boolean
 
Private Sub Class_Initialize()
    Call StopWatchInit
End Sub
 
'*** 初期化 ***
Private Function StopWatchInit()
    glStartTime = 0
    glIntervalTime = 0
    gbIsTimerStarted = False
End Function
 
'*** 測定開始 ***
Public Sub StartT()
    Debug.Assert gbIsTimerStarted = False
    glStartTime = timeGetTime()
    glIntervalTime = glStartTime
    gbIsTimerStarted = True
End Sub
 
'*** 測定停止 ***
Public Function StopT() As Long
    Debug.Assert gbIsTimerStarted = True
    StopT = timeGetTime() - glStartTime
    Call StopWatchInit
End Function
 
'*** 開始から現在までの総経過時間 ***
Public Function LapTime() As Long
    Debug.Assert gbIsTimerStarted = True
    LapTime = timeGetTime() - glStartTime
End Function
 
'*** 時間間隔 ***
Public Function IntervalTime() As Long
    Dim lCurTime As Long
    Debug.Assert gbIsTimerStarted = True
    lCurTime = timeGetTime()
    IntervalTime = lCurTime - glIntervalTime
    glIntervalTime = lCurTime
End Function
 
Private Sub Class_Terminate()
    Call StopWatchInit
End Sub
