VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExecAddInMacro 
   Caption         =   "アドインマクロ実行"
   ClientHeight    =   11550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5580
   OleObjectBlob   =   "ExecAddInMacro.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ExecAddInMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' execute addin macro v1.0

Const sMACRO_NAME As String = "アドインマクロ実行"
Const sEXECADDINMACRO_MACRONAME As String = ""

Private Sub ExecButton_Click()
    Dim sExecAddinMacroName As String
    sExecAddinMacroName = Me.MacroNamesListBox.Value
    
    'アドインマクロ名 前回値保存
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.WriteItemToFile(sSettingFilePath, "sEXECADDINMACRO_MACRONAME", sExecAddinMacroName)
    
    'MsgBox Me.MacroNamesListBox.Value, vbOKOnly, sMACRO_NAME
    Application.Run sExecAddinMacroName
    Unload Me
End Sub

Private Sub CancelButton_Click()
    MsgBox "キャンセルされたため、処理を中断します", vbOKOnly, sMACRO_NAME
    Unload Me
End Sub

Private Sub MacroNamesListBox_Click()
    'MsgBox Me.MacroNamesListBox.Value, vbOKOnly, sMACRO_NAME
    'Application.Run Me.MacroNamesListBox.Value
    'Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim vProcNames As Variant
    Set vProcNames = CreateObject("System.Collections.ArrayList")
    
    'アドインマクロ名 前回値取得
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    Dim sExecAddinMacroName As String
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sEXECADDINMACRO_MACRONAME", sExecAddinMacroName, sEXECADDINMACRO_MACRONAME, False)
    
    'アドインマクロ名一覧取得
    Call ExtractPublicSubMacros("Macros", ThisWorkbook, vProcNames)
    
    'アドインマクロ名に該当するインデックス取得
    Dim lSelectIdx As Long
    lSelectIdx = 0
    If sExecAddinMacroName = "" Then
        'Do Nothing
    Else
        Dim vProcName As Variant
        For Each vProcName In vProcNames
            If vProcName = sExecAddinMacroName Then
                Exit For
            Else
                lSelectIdx = lSelectIdx + 1
            End If
        Next
    End If
    
    'リストボックス表示
    With MacroNamesListBox
        '.Height = 9 * vProcNames.Count
        For Each vProcName In vProcNames
            .AddItem vProcName
        Next
        .ListIndex = lSelectIdx
        .SetFocus
    End With
End Sub

' ==================================================================
' = 概要    公開マクロ名を取得する
' = 引数    wTrgtBook   Workbook    [in]    抽出対象ブック名
' = 引数    sModuleName String      [in]    抽出対象モジュール名
' = 引数    vProcNames  Variant　   [out]   公開マクロ名
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    ExecMacro.frm
' ==================================================================
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


