VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectActivationSheet 
   Caption         =   "�V�[�g�I���E�B���h�E"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SelectActivationSheet.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SelectActivationSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' select activation sheet macro v1.1

'�������ݒ� �������灥����
Const lFORM_HEIGHT_MERGIN As Long = 30
Const lFONT_SIZE As Long = 10
Const sFONT_NAME As String = "�l�r �S�V�b�N"
'�������ݒ� �������灣����

Const sMACRO_NAME As String = "�V�[�g�I���E�B���h�E�\��"

Private Sub SheetNamesListBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Const lKEY_ENTER As Long = 13
    Const lKEY_ESC As Long = 27
    If KeyAscii = lKEY_ENTER Then
        '�V�[�g�A�N�e�B�u��
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
    '�V�[�g�A�N�e�B�u��
    Dim sSheetName As String
    sSheetName = Me.SheetNamesListBox.Value
    ActiveWorkbook.Sheets(sSheetName).Activate
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.Height = Application.Height - lFORM_HEIGHT_MERGIN
    
    '���X�g�{�b�N�X�\��
    With SheetNamesListBox
        Debug.Print Me.Height
        .Height = Me.Height - 70
        .Font.Size = lFONT_SIZE
        .Font = sFONT_NAME
        
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
    
    '�t�H�[����Excel�E�B���h�E�̒����ɕ\��(�f���A���f�B�X�v���C�΍�)
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ((Application.Height - Me.Height) / 2)
    Me.Left = Application.Left + ((Application.Width - Me.Width) / 2)
End Sub

