VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectActivationSheet 
   Caption         =   "シート選択ウィンドウ"
   ClientHeight    =   9165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SelectActivationSheet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SelectActivationSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' select activation sheet macro v1.0

Const sMACRO_NAME As String = "シート選択ウィンドウ表示"

Const lKEY_ENTER As Long = 13
Const lKEY_ESC As Long = 27

Private Sub SheetNamesListBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = lKEY_ENTER Then
        'シートアクティブ化
        Dim sSheetName As String
        sSheetName = Me.SheetNamesListBox.Value
        ActiveWorkbook.Sheets(sSheetName).Activate
        Unload Me
    ElseIf KeyAscii = lKEY_ESC Then
        Unload Me
    Else
        'Do Nothing
    End If
End Sub

Private Sub SheetNamesListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'シートアクティブ化
    Dim sSheetName As String
    sSheetName = Me.SheetNamesListBox.Value
    ActiveWorkbook.Sheets(sSheetName).Activate
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    'リストボックス表示
    With SheetNamesListBox
        Dim lCurSheetIdx As Long
        Dim lSheetIdx As Long
        lSheetIdx = 0
        Dim oSheet As Worksheet
        For Each oSheet In ActiveWorkbook.Sheets
            If oSheet.Visible = True Then
                .AddItem oSheet.Name
                If ActiveSheet.Name = oSheet.Name Then
                    lCurSheetIdx = lSheetIdx
                End If
                lSheetIdx = lSheetIdx + 1
            End If
        Next
        .ListIndex = lCurSheetIdx
        .SetFocus
    End With
    
    'フォームをExcelウィンドウの中央に表示(デュアルディスプレイ対策)
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ((Application.Height - Me.Height) / 2)
    Me.Left = Application.Left + ((Application.Width - Me.Width) / 2)
End Sub

