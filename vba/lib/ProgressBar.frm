VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "モードレス表示を使用した進捗表示"
   ClientHeight    =   2316
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   2580
   OleObjectBlob   =   "ProgressBar_v11.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'======================================================
' 設定値
'======================================================
Private Const REPAINT_TIME As Double = 0.1 '[s]

Private Const LEFT_OFFSET As Long = 10
Private Const HEIGHT_BAR As Long = 30
Private Const HEIGHT_SPACE As Long = 10
Private Const HEIGHT_CANCELBUTTON As Long = 25
Private Const WIDTH_CANCELBUTTON As Long = 90
Private Const WIDTH_WINDOW As Long = 350

Private Const BAR_COLOR_R As Long = 248
Private Const BAR_COLOR_G As Long = 150
Private Const BAR_COLOR_B As Long = 150
Private Const FONT_NAME As String = "MS ゴシック"
Private Const FONT_SIZE_LABEL As Long = 14
Private Const FONT_SIZE_BAR As Long = 15
Private Const FONT_SIZE_CANCELBUTTON As Long = 12

'======================================================
' 定数＆変数
'======================================================
Private Const HEIGHT_WINDOWTITLE As Long = 20

Private glBarMaxWidth As Long
Private gdOldTime As Double
Private glProgMsgLineNum As Long
Private gbIsCanceled As Boolean

'======================================================
' 本処理
'======================================================
Private Sub UserForm_Initialize()
    With Me
        .Caption = "進捗状況"
        
        With .ProgMsg
            .Caption = ""
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE_LABEL
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
        With .CancelButton
            .Caption = "Cancel"
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE_CANCELBUTTON
            '.TextAlign = fmTextAlignCenter
        End With
    End With
    
    glProgMsgLineNum = 0
    
    Call FormResize
    
    glBarMaxWidth = Me.ProgBarFrame.Width - 6
    gdOldTime = Timer
    gbIsCanceled = False
End Sub

Private Sub UserForm_Terminate()
    'Do Nothing
End Sub

Private Sub CancelButton_Click()
    gbIsCanceled = True
End Sub

Public Property Get IsCanceled()
    IsCanceled = gbIsCanceled
End Property

Public Property Let Title( _
    ByVal sTitle As String _
)
    Me.Caption = sTitle
End Property

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
    
    'キャプション設定
    With Me
        .ProgMsg.Caption = sProgMsg
        .ProgPer.Caption = Int(dProgPer * 100) & " [%] 完了..."
        .ProgBar.Width = glBarMaxWidth * dProgPer 'プログレスバーの進捗表示を更新
    End With
    
    '再描画
    Dim dCurTime As Double
    dCurTime = Timer
    If (dCurTime - gdOldTime) > REPAINT_TIME Then
        DoEvents
        gdOldTime = dCurTime
    End If
End Function

Private Function FormResize()
    With Me
        With .ProgMsg
            .Top = HEIGHT_SPACE
            .Left = LEFT_OFFSET
            .Width = Me.Width - (.Left * 2)
            .Height = .Font.Size * glProgMsgLineNum
        End With
        With .ProgBarFrame
            If glProgMsgLineNum = 0 Then
                .Top = HEIGHT_SPACE
            Else
                .Top = HEIGHT_SPACE + Me.ProgMsg.Height + HEIGHT_SPACE
            End If
            .Left = LEFT_OFFSET
            .Width = WIDTH_WINDOW - (.Left * 2)
            .Height = HEIGHT_BAR
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
        With .CancelButton
            .Width = WIDTH_CANCELBUTTON
            .Height = HEIGHT_CANCELBUTTON
            .Top = Me.ProgBarFrame.Top + Me.ProgBarFrame.Height + HEIGHT_SPACE
            .Left = (WIDTH_WINDOW - .Width) / 2
            .SetFocus
        End With
        
        .Width = WIDTH_WINDOW
        .Height = .CancelButton.Top + .CancelButton.Height + HEIGHT_SPACE + HEIGHT_WINDOWTITLE
        .Top = (Application.Height - .Height) / 2
        .Left = (Application.Width - .Width) / 2
    End With
End Function

