VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetVisibleSetting 
   Caption         =   "シート表示・非表示切り替え"
   ClientHeight    =   4476
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6000
   OleObjectBlob   =   "SheetVisibleSetting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SheetVisibleSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' sheet visible setting v1.0

Const CHK_BOX_HEIGHT = 15
Const CHK_BOX_WIDTH = 300
Const CHK_BOX_INTERVAL = 0
Const CHK_BOX_LEFT = 5

Private Sub UserForm_Initialize()
    Dim lShtCnt As Long
    Dim lShtNum As Long
    Dim lChkBoxTotalHeight As Long
    Dim myCheckBox As Control
    
    lShtNum = ActiveWorkbook.Sheets.Count
    With SheetSelFrame
        lChkBoxTotalHeight = (CHK_BOX_HEIGHT + CHK_BOX_INTERVAL) * lShtNum
        If .Height < lChkBoxTotalHeight Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = lChkBoxTotalHeight
        Else
            'Do Nothing
        End If
        '=== チェックボックス表示 ===
        For lShtCnt = 1 To lShtNum
            Set myCheckBox = .Controls.Add("Forms.CheckBox.1")
            With myCheckBox
                .Height = CHK_BOX_HEIGHT
                .Width = CHK_BOX_WIDTH
                .Left = CHK_BOX_LEFT
                .Top = (CHK_BOX_HEIGHT + CHK_BOX_INTERVAL) * (lShtCnt - 1)
                .Caption = ActiveWorkbook.Sheets(lShtCnt).Name
            End With
        Next lShtCnt
    End With
    '=== 「全て選択」にフォーカスする ===
    With ChkBox_AllSelect
        .SetFocus
    End With
End Sub

Private Sub ChkBox_AllSelect_Click()
    Dim lShtCnt As Long
    Dim bSetCnt As Boolean
    bSetCnt = ChkBox_AllSelect.Value
    For lShtCnt = 0 To ActiveWorkbook.Sheets.Count - 1
        SheetSelFrame.Controls.Item(lShtCnt) = bSetCnt
    Next lShtCnt
End Sub

Private Sub HiddenButton_Click()
    Dim lShtCnt As Long
    Dim bIsExistShowSht As Boolean
    Dim bIsExistCheck As Boolean
    Dim abIsVisible() As Boolean
    
    With ActiveWorkbook
        '=== 表示/非表示チェック ===
        bIsExistCheck = False
        bIsExistShowSht = False
        ReDim Preserve abIsVisible(.Sheets.Count - 1)
        For lShtCnt = 1 To .Sheets.Count
            If SheetSelFrame.Controls.Item(lShtCnt - 1) = True Then
                abIsVisible(lShtCnt - 1) = False
                bIsExistCheck = True
            Else
                If .Sheets(lShtCnt).Visible = True Then
                    abIsVisible(lShtCnt - 1) = True
                    bIsExistShowSht = True
                Else
                    abIsVisible(lShtCnt - 1) = False
                End If
            End If
        Next lShtCnt
        
        '=== 表示/非表示切り替え ===
        If bIsExistCheck = True Then
            If bIsExistShowSht = True Then
                For lShtCnt = 1 To .Sheets.Count
                    .Sheets(lShtCnt).Visible = abIsVisible(lShtCnt - 1)
                Next lShtCnt
                Unload Me
            Else
                MsgBox "全てを非表示にできません！"
            End If
        Else
            MsgBox "一つもチェックされていません！"
        End If
    End With
End Sub

Private Sub ShowButton_Click()
    Dim lShtCnt As Long
    Dim bIsExistCheck As Boolean
    Dim abIsVisible() As Boolean
    
    With ActiveWorkbook
        '=== 表示/非表示チェック ===
        bIsExistCheck = False
        ReDim Preserve abIsVisible(.Sheets.Count - 1)
        For lShtCnt = 1 To .Sheets.Count
            If SheetSelFrame.Controls.Item(lShtCnt - 1) = True Then
                abIsVisible(lShtCnt - 1) = True
                bIsExistCheck = True
            Else
                If .Sheets(lShtCnt).Visible = True Then
                    abIsVisible(lShtCnt - 1) = True
                Else
                    abIsVisible(lShtCnt - 1) = False
                End If
            End If
        Next lShtCnt
        
        '=== 表示/非表示切り替え ===
        If bIsExistCheck = True Then
            For lShtCnt = 1 To .Sheets.Count
                .Sheets(lShtCnt).Visible = abIsVisible(lShtCnt - 1)
            Next lShtCnt
            Unload Me
        Else
            MsgBox "一つもチェックされていません！"
        End If
    End With
End Sub

