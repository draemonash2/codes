VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExecAddInMacro 
   Caption         =   "�A�h�C���}�N�����s"
   ClientHeight    =   11550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5580
   OleObjectBlob   =   "ExecAddInMacro.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ExecAddInMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' execute addin macro v1.0

Const sMACRO_NAME As String = "�A�h�C���}�N�����s"
Const sEXECADDINMACRO_MACRONAME As String = ""

Private Sub ExecButton_Click()
    Dim sExecAddinMacroName As String
    sExecAddinMacroName = Me.MacroNamesListBox.Value
    
    '�A�h�C���}�N���� �O��l�ۑ�
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.WriteItemToFile(sSettingFilePath, "sEXECADDINMACRO_MACRONAME", sExecAddinMacroName)
    
    'MsgBox Me.MacroNamesListBox.Value, vbOKOnly, sMACRO_NAME
    Application.Run sExecAddinMacroName
    Unload Me
End Sub

Private Sub CancelButton_Click()
    MsgBox "�L�����Z�����ꂽ���߁A�����𒆒f���܂�", vbOKOnly, sMACRO_NAME
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
    
    '�A�h�C���}�N���� �O��l�擾
    Dim clSetting As New SettingFile
    Dim sSettingFilePath As String
    Dim sExecAddinMacroName As String
    sSettingFilePath = GetAddinSettingFilePath()
    Call clSetting.ReadItemFromFile(sSettingFilePath, "sEXECADDINMACRO_MACRONAME", sExecAddinMacroName, sEXECADDINMACRO_MACRONAME, False)
    
    '�A�h�C���}�N�����ꗗ�擾
    Call ExtractPublicSubMacros("Macros", ThisWorkbook, vProcNames)
    
    '�A�h�C���}�N�����ɊY������C���f�b�N�X�擾
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
    
    '���X�g�{�b�N�X�\��
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
' = �T�v    ���J�}�N�������擾����
' = ����    wTrgtBook   Workbook    [in]    ���o�Ώۃu�b�N��
' = ����    sModuleName String      [in]    ���o�Ώۃ��W���[����
' = ����    vProcNames  Variant�@   [out]   ���J�}�N����
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    ExecMacro.frm
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
            sSearchPattern = "^ *(Private|Public)* *(Sub|Function)+ +([��-Ꞃ�-��[�@-���`-�y��-���O-�X\w]+)\("
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


