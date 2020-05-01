VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
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
'        Sleep 1000 'ms �P��
'        Debug.Print "Lap time is " & clStopWatch.LapTime
'        Debug.Print "Interval time is " & clStopWatch.IntervalTime
'
'        Sleep 1500 'ms �P��
'        Debug.Print "Lap time is " & clStopWatch.LapTime
'        Debug.Print "Interval time is " & clStopWatch.IntervalTime
'
'        Sleep 2000 'ms �P��
'        Debug.Print "Lap time is " & clStopWatch.LapTime
'        Debug.Print "Interval time is " & clStopWatch.IntervalTime
'
'        Sleep 2500 'ms �P��
'        Debug.Print clStopWatch.StopT
'    End Sub

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
 
Dim glStartTime As Long
Dim glIntervalTime As Long
Dim gbIsTimerStarted As Boolean
 
Private Sub Class_Initialize()
    Call StopWatchInit
End Sub
 
'*** ������ ***
Private Function StopWatchInit()
    glStartTime = 0
    glIntervalTime = 0
    gbIsTimerStarted = False
End Function
 
'*** ����J�n ***
Public Sub StartT()
    Debug.Assert gbIsTimerStarted = False
    glStartTime = timeGetTime()
    glIntervalTime = glStartTime
    gbIsTimerStarted = True
End Sub
 
'*** �����~ ***
Public Function StopT() As Long
    Debug.Assert gbIsTimerStarted = True
    StopT = timeGetTime() - glStartTime
    Call StopWatchInit
End Function
 
'*** �J�n���猻�݂܂ł̑��o�ߎ��� ***
Public Function LapTime() As Long
    Debug.Assert gbIsTimerStarted = True
    LapTime = timeGetTime() - glStartTime
End Function
 
'*** ���ԊԊu ***
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