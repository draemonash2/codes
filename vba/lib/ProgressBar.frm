VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "モードレス表示を使用した進捗表示"
   ClientHeight    =   2544
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

'ProgressBar v1.3
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
Private Const FONT_SIZE_BAR As Long = 15
Private Const FONT_SIZE_BUTTON As Long = 12

'======================================================
' 定数＆変数
'======================================================
Private Const HEIGHT_WINDOWTITLE As Long = 20

Private glBarMaxWidth As Long
Private gdOldTime As Double
Private glProgMsgLineNum As Long
Private gdStartTime As Double
Private gbIsCanceled As Boolean
Private gbIsSuspended As Boolean

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
    gbIsCanceled = False
    gbIsSuspended = False
End Sub

Private Sub UserForm_Terminate()
    'Do Nothing
End Sub

Public Property Get IsCanceled()
    IsCanceled = gbIsCanceled
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
    Dim dNow As Double
    Dim lElapsedTime As Long
    dNow = Timer
    If dNow - gdStartTime > 0 Then
        lElapsedTime = dNow - gdStartTime
    Else
        lElapsedTime = ((60 * 60 * 24) - gdStartTime) + dNow
    End If
    
    'キャプション設定
    With Me
        .ProgMsg.Caption = sProgMsg
        .ElpsdTime.Caption = "経過時間：" & lElapsedTime & " [秒]"
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


