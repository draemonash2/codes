VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "モードレス表示を使用した進捗表示"
   ClientHeight    =   3048
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3888
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'ProgressBar v1.5
'
'<<Usage Sample>>
'    Sub test()
'        Dim oProgBar As New ProgressBar
'
'        'プログレスバー読込＆表示
'        Load oProgBar
'        oProgBar.Show vbModeless
'
'        Dim lIdx As Long
'        Dim lMax As Long
'        lMax = 10
'        For lIdx = 0 To lMax
'            oProgBar.Update (lIdx / lMax) '0～1を指定
'            If oProgBar.IsCanceled = True Then
'                Exit For
'            End If
'            Application.Wait Now + TimeValue("00:00:01")
'        Next lIdx
'
'        'プログレスバー終了
'        oProgBar.Hide
'        Unload oProgBar
'    End Sub

'======================================================
' 設定値
'======================================================
Private Const REPAINT_TIME As Double = 0.1 '[s]

Private Const LEFT_OFFSET As Long = 10
Private Const HEIGHT_BAR As Long = 30
Private Const HEIGHT_SPACE As Long = 10
Private Const HEIGHT_BUTTON As Long = 25
Private Const WIDTH_BUTTON As Long = 90
Private Const WIDTH_WINDOW As Long = 350
Private Const BUTTON_SPACE As Long = 40

Private Const BAR_COLOR_R As Long = 248
Private Const BAR_COLOR_G As Long = 150
Private Const BAR_COLOR_B As Long = 150
Private Const FONT_NAME As String = "MS ゴシック"
Private Const FONT_SIZE_LABEL As Long = 14
Private Const FONT_SIZE_ELPSDTIME As Long = 12
Private Const FONT_SIZE_REMTIME As Long = 12
Private Const FONT_SIZE_BAR As Long = 15
Private Const FONT_SIZE_BUTTON As Long = 12

Private Const ELAPSED_TIME_DIFF_COUNT_MAX As Long = 10

'======================================================
' 定数＆変数
'======================================================
Private Const DATE_SECOND = 86400 '60[s] * 60[m] * 24[h]
Private Const HEIGHT_WINDOWTITLE As Long = 20

Private glBarMaxWidth As Long
Private gdOldTime As Double
Private glProgMsgLineNum As Long
Private gdStartTime As Double
Private gsStartDate As String
Private gbIsCanceled As Boolean
Private gbIsSuspended As Boolean
Private gdElapsedTime As Double
Private gdProgPerLastCalc As Double
Private glElapsedTimeStoreNum As Long
Private gdElapsedTimeDiffTable() As Double
Private gdElapsedTimeLastCalc As Double
Private gdRemainingTime As Double

'======================================================
' 本処理
'======================================================
Private Sub CancelButton_Click()
    gbIsCanceled = True
End Sub

Private Sub SuspendButton_Click()
    If gbIsSuspended = True Then
        gbIsSuspended = False
    Else
        gbIsSuspended = True
    End If
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Caption = "進捗状況"
        
        With .ProgMsg
            .Caption = ""
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE_LABEL
        End With
        With .ElpsdTime
            .Caption = ""
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE_ELPSDTIME
        End With
        With .RemTime
            .Caption = ""
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE_REMTIME
        End With
        With .ProgBarFrame
            .Caption = ""
        End With
        With .ProgBar
            .Caption = ""
            .BackColor = RGB(BAR_COLOR_R, BAR_COLOR_G, BAR_COLOR_B)
        End With
        With .ProgPer
            .Caption = ""
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE_BAR
            .BackStyle = fmBackStyleTransparent
            .TextAlign = fmTextAlignCenter
        End With
        With .SuspendButton
            .Caption = "一時停止"
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE_BUTTON
        End With
        With .CancelButton
            .Caption = "キャンセル"
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE_BUTTON
            '.SetFocus
        End With
    End With
    
    glProgMsgLineNum = 0
    
    Call FormResize
    
    glBarMaxWidth = Me.ProgBarFrame.Width - 6
    gdOldTime = Timer
    gdStartTime = Timer
    gsStartDate = Date
    gbIsCanceled = False
    gbIsSuspended = False
    gdElapsedTime = 0
    gdProgPerLastCalc = 0
    glElapsedTimeStoreNum = 0
    ReDim Preserve gdElapsedTimeDiffTable(ELAPSED_TIME_DIFF_COUNT_MAX - 1)
    Dim i
    For i = 0 To ELAPSED_TIME_DIFF_COUNT_MAX - 1
        gdElapsedTimeDiffTable(i) = 0
    Next i
    gdElapsedTimeLastCalc = 0
    gdRemainingTime = 7200
End Sub

Private Sub UserForm_Terminate()
    'Do Nothing
End Sub

Public Property Get IsCanceled()
    IsCanceled = gbIsCanceled
End Property

Public Property Get ElapsedTime()
    ElapsedTime = gdElapsedTime
End Property

Public Property Get RemainingTime()
    RemainingTime = gdRemainingTime
End Property

'Public Property Get IsSuspended()
'    IsSuspended = gbIsSuspended
'End Property

Public Property Let Title( _
    ByVal sTitle As String _
)
    Me.Caption = sTitle
End Property

'覚書：Update() 関数一回呼び出しにかかる処理時間＝150[us]
'      すなわち、10000回連続呼び出しにかかる時間＝2[s]
Public Function Update( _
    ByVal dProgPer As Double, _
    Optional ByVal sProgMsg As String _
)
    Debug.Assert 0 <= dProgPer And dProgPer <= 1
    
    '行数算出
    If sProgMsg = "" Then
        glProgMsgLineNum = 0
    Else
        glProgMsgLineNum = (Len(sProgMsg) - Len(Replace(sProgMsg, vbNewLine, ""))) / 2 + 1
    End If
    Call FormResize
    
    '経過時間算出
    Dim sDateOld As String
    Dim sDateNow As String
    Dim dSecondOld As Double
    Dim dSecondNow As Double
    Dim lDateDiff As Long
    sDateOld = gsStartDate
    sDateNow = Date
    dSecondOld = gdStartTime
    dSecondNow = Timer
    lDateDiff = DateDiff("d", sDateOld, sDateNow)
    If lDateDiff > 0 Then
        gdElapsedTime = (DATE_SECOND * (lDateDiff - 1)) + (DATE_SECOND - dSecondOld) + dSecondNow
    ElseIf lDateDiff = 0 Then
        gdElapsedTime = dSecondNow - dSecondOld
    Else
        gdElapsedTime = 0
    End If
    
    '残り時間算出
    Dim dProgPerCur As Double
    Dim dProgPerDiff As Double
    Dim dElapsedTimeCur As Double
    dProgPerCur = dProgPer
    dProgPerDiff = dProgPerCur - gdProgPerLastCalc
    dElapsedTimeCur = gdElapsedTime
    If Int(dProgPerDiff * 100) >= 1 Then
        Dim dProgPerRem As Double
        Dim dElapsedTimeDiff As Double
        Dim dElapsedTime1PerCur As Double
        dProgPerRem = 100 - (dProgPerCur * 100)
        dElapsedTimeDiff = dElapsedTimeCur - gdElapsedTimeLastCalc
        dElapsedTime1PerCur = dElapsedTimeDiff / Int(dProgPerDiff * 100)
        Dim dElapsedTimeSum As Double
        dElapsedTimeSum = dElapsedTime1PerCur
        Dim i As Long
        For i = 0 To glElapsedTimeStoreNum - 1 Step 1
            dElapsedTimeSum = dElapsedTimeSum + gdElapsedTimeDiffTable(i)
        Next i
        Dim dElapsedTimeAvg As Double
        dElapsedTimeAvg = dElapsedTimeSum / (glElapsedTimeStoreNum + 1)
        gdRemainingTime = dElapsedTimeAvg * dProgPerRem
        
        For i = ELAPSED_TIME_DIFF_COUNT_MAX - 1 To 1 Step -1
            gdElapsedTimeDiffTable(i) = gdElapsedTimeDiffTable(i - 1)
        Next i
        gdElapsedTimeDiffTable(0) = dElapsedTime1PerCur
        If glElapsedTimeStoreNum < ELAPSED_TIME_DIFF_COUNT_MAX Then
            glElapsedTimeStoreNum = glElapsedTimeStoreNum + 1
        Else
            'Do Nothing
        End If
        gdProgPerLastCalc = dProgPerCur
        gdElapsedTimeLastCalc = dElapsedTimeCur
    Else
        'Do Nothing
    End If
    
    'キャプション設定
    With Me
        .ProgMsg.Caption = sProgMsg
        .ElpsdTime.Caption = "経過時間：" & ConvSec2SplitTime(Int(gdElapsedTime))
        .RemTime.Caption = "残り時間：" & ConvSec2SplitTime(Application.RoundUp(gdRemainingTime, 0))
        .ProgPer.Caption = Int(dProgPer * 100) & " [%]"
        .ProgBar.Width = glBarMaxWidth * dProgPer 'プログレスバーの進捗表示を更新
    End With
    
    '再描画
    Dim dCurTime As Double
    dCurTime = Timer
    If (dCurTime - gdOldTime) > REPAINT_TIME Then
        DoEvents
        gdOldTime = dCurTime
    End If
    
    '「一時停止」ボタン押下受付
    Do While gbIsSuspended = True
        DoEvents
        If gbIsCanceled = True Then
            Exit Do
        Else
            'Do Nothing
        End If
    Loop
End Function

Private Function FormResize()
    Dim lHeightOffset As Long
    lHeightOffset = 0
    With Me
        With .ProgMsg
            .Top = lHeightOffset + HEIGHT_SPACE
            .Left = LEFT_OFFSET
            .Width = WIDTH_WINDOW - (.Left * 2)
            .Height = .Font.Size * glProgMsgLineNum
            If glProgMsgLineNum = 0 Then
                'Do Nothing
            Else
                lHeightOffset = .Top + .Height
            End If
        End With
        With .ElpsdTime
            .Top = lHeightOffset + HEIGHT_SPACE
            .Left = LEFT_OFFSET
            .Width = WIDTH_WINDOW - (.Left * 2)
            .Height = .Font.Size
            lHeightOffset = .Top + .Height
        End With
        With .RemTime
            .Top = lHeightOffset + HEIGHT_SPACE
            .Left = LEFT_OFFSET
            .Width = WIDTH_WINDOW - (.Left * 2)
            .Height = .Font.Size
            lHeightOffset = .Top + .Height
        End With
        With .ProgBarFrame
            .Top = lHeightOffset + HEIGHT_SPACE
            .Left = LEFT_OFFSET
            .Width = WIDTH_WINDOW - (.Left * 2)
            .Height = HEIGHT_BAR
            lHeightOffset = .Top + .Height
        End With
        With .ProgBar 'Top,Left は .ProgBarFrame からのオフセット
            .Top = 1
            .Left = 1
            .Width = 0
            .Height = Me.ProgBarFrame.Height - 6
        End With
        With .ProgPer 'Top,Left は .ProgBarFrame からのオフセット
            .Width = Me.ProgBarFrame.Width - 6
            .Height = .Font.Size
            .Top = (Me.ProgBarFrame.Height - .Height) / 2 - 2
            .Left = Me.ProgBar.Left
        End With
        With .SuspendButton
            .Width = WIDTH_BUTTON
            .Height = HEIGHT_BUTTON
            .Top = lHeightOffset + HEIGHT_SPACE
            .Left = (WIDTH_WINDOW - (WIDTH_BUTTON * 2 + BUTTON_SPACE)) / 2
            'lHeightOffset = .Top + .Height '.CancelButton にて加算
        End With
        With .CancelButton
            .Width = WIDTH_BUTTON
            .Height = HEIGHT_BUTTON
            .Top = lHeightOffset + HEIGHT_SPACE
            .Left = Me.SuspendButton.Left + Me.SuspendButton.Width + BUTTON_SPACE
            lHeightOffset = .Top + .Height
        End With
        
        .Width = WIDTH_WINDOW
        .Height = lHeightOffset + HEIGHT_SPACE + HEIGHT_WINDOWTITLE
        .Top = (Application.Height - .Height) / 2
        .Left = (Application.Width - .Width) / 2
    End With
End Function

Private Function ConvSec2SplitTime( _
    ByVal lRawSec As Long _
)
    Dim lOutSec As Long
    Dim lOutMin As Long
    Dim lOutHour As Long
    Dim lMod As Long
    
    lMod = lRawSec
    lOutHour = Fix(lMod / (60 * 60))
    
    lMod = Fix(lMod - (lOutHour * (60 * 60)))
    lOutMin = Fix(lMod / 60)
    
    lMod = Fix(lMod - (lOutMin * 60))
    lOutSec = lMod
    
    If lOutHour = 0 Then
        If lOutMin = 0 Then
            ConvSec2SplitTime = lOutSec & " 秒"
        Else
            ConvSec2SplitTime = lOutMin & " 分 " & lOutSec & " 秒"
        End If
    Else
        ConvSec2SplitTime = lOutHour & " 時間 " & lOutMin & " 分 " & lOutSec & " 秒"
    End If
End Function

