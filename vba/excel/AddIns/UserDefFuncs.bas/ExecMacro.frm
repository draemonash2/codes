VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExecMacro 
   Caption         =   "アドインマクロ実行"
   ClientHeight    =   11550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5580
   OleObjectBlob   =   "ExecMacro.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ExecMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' execute addin macro v1.0
'★関数ヘッダ
'★アドイン名記憶

Private Sub CancelButton_Click()
    MsgBox "キャンセルされたため、処理を中断します"
    Unload Me
End Sub

Private Sub ExecButton_Click()
    'MsgBox Me.MacroNamesListBox.Value
    Application.Run Me.MacroNamesListBox.Value
    Unload Me
End Sub

Private Sub MacroNamesListBox_Click()
    'MsgBox Me.MacroNamesListBox.Value
    'Application.Run Me.MacroNamesListBox.Value
    'Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim vProcNames As Variant
    Set vProcNames = CreateObject("System.Collections.ArrayList")
    
    'アドインマクロ名取得
    Call ExtractPublicSubMacros("Macros", ThisWorkbook, vProcNames)
    
    'アドインマクロ名表示
    With MacroNamesListBox
        Dim vProcName As Variant
        For Each vProcName In vProcNames
            .AddItem vProcName
        Next
        .SetFocus
    End With
End Sub

Private Function ExtractPublicSubMacros( _
    ByVal sModuleName As String, _
    ByRef wTrgtBook As Workbook, _
    ByRef vProcNames As Variant _
)
    Dim oRegExp As Object
    Set oRegExp = CreateObject("VBScript.RegExp")
    With wTrgtBook.VBProject.VBComponents(sModuleName).CodeModule
        oRegExp.IgnoreCase = True
        oRegExp.Global = True
        Dim lLineIdx As Long
        For lLineIdx = 1 To .CountOfLines
            Dim sTargetStr As String
            Dim sSearchPattern As String
            sTargetStr = .Lines(lLineIdx, 1)
            sSearchPattern = "^ *(Private|Public)* *(Sub|Function)+ +([一-龠ぁ-んーァ-ヶＡ-Ｚａ-ｚ０-９\w]+)\("
            oRegExp.Pattern = sSearchPattern
            Dim oMatchResult As Object
            Set oMatchResult = oRegExp.Execute(sTargetStr)
            If oMatchResult.Count = 0 Then
                'Do Nothing
            Else
                If oMatchResult(0).SubMatches(0) = "Public" Or _
                    oMatchResult(0).SubMatches(0) = "" Then
                    If oMatchResult(0).SubMatches(1) = "Sub" Then
                        vProcNames.Add oMatchResult(0).SubMatches(2)
                    Else
                        'Do Nothing
                    End If
                Else
                    'Do Nothing
                End If
            End If
        Next lLineIdx
    End With
End Function


