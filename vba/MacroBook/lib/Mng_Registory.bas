Attribute VB_Name = "Mng_Registory"
Option Explicit

' registory manage library v1.01

'「REG_MULTI_SZ」と「REG_EXPAND_SZ」は非対応
Public Enum E_REG_ENUM
    REG_SZ = 0
    REG_DWORD
    REG_BINARY
End Enum

Public Enum E_REG_OPERATION
    REG_ADDMOD = 0
    REG_DELETE
End Enum

Public Type T_REG_VALUES
    sName As String
    sData As String
    eType As E_REG_ENUM
    eOperation As E_REG_OPERATION
End Type

Public Type T_REG_KEYS
    sKey As String
    eOperation As E_REG_OPERATION
    atRegValues() As T_REG_VALUES
End Type

Public Type T_REG_STRUCT
    atRegKeys() As T_REG_KEYS
End Type

Private gbIsRegWrite As Boolean

'********************************************************************************
'* 外部関数定義
'********************************************************************************
' ==================================================================
' = 概要    レジストリ変更有効化
' = 引数    なし
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_Registory.bas
' ==================================================================
Public Sub EnableRegWrite()
    gbIsRegWrite = True
End Sub

' ==================================================================
' = 概要    レジストリ変更無効化
' = 引数    なし
' = 戻値    なし
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_Registory.bas
' ==================================================================
Public Sub DisableRegWrite()
    gbIsRegWrite = False
End Sub

' ==================================================================
' = 概要    レジストリ変更
' = 引数    sRegFileTitle   [in]    String          ファイル名
' = 引数    tRegStruct      [in]    T_REG_STRUCT    値、キー
' = 戻値    なし
' = 覚書    なし
' = 依存    Mng_Registory.bas/GetRegKeysText()
' = 所属    Mng_Registory.bas
' ==================================================================
Public Sub SetRegistry( _
    ByVal sRegFileTitle As String, _
    ByRef tRegStruct As T_REG_STRUCT _
)
    On Error Resume Next
    Dim sRegFilePath As String
    sRegFilePath = Environ("tmp") & "\" & sRegFileTitle & ".reg"
    Debug.Print sRegFilePath
    Open sRegFilePath For Output As #1
    
    Dim sOutText As String
    sOutText = ""
    sOutText = sOutText & "Windows Registry Editor Version 5.00" & vbCrLf & vbCrLf
    sOutText = sOutText & GetRegKeysText(tRegStruct.atRegKeys)
    
    Print #1, sOutText
    Close #1
    
    If gbIsRegWrite = True Then
        Dim sCommand As String
        sCommand = "cmd.exe /c """ & sRegFilePath & """"
        Shell sCommand, vbMinimizedFocus
    Else
        'Do Nothing
    End If
    
    On Error GoTo 0
End Sub
    Private Sub Test_SetRegistry()
        Dim sRegFilePath As String
        sRegFilePath = "test2"
        
        Dim tRegStruct As T_REG_STRUCT
        With tRegStruct
            ReDim Preserve .atRegKeys(3)
            With .atRegKeys(0)
                .sKey = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\" & Application.Version & "\Excel\DisabledShortcutKeysCheckBoxes"
                .eOperation = REG_ADDMOD
                ReDim Preserve .atRegValues(1)
                With .atRegValues(0)
                    .sName = "F1key"
                    .sData = 1
                    .eType = REG_SZ
                    .eOperation = REG_ADDMOD
                End With
                With .atRegValues(1)
                    .sName = "F1key2"
                    .sData = "c:\test\test.txt"
                    .eType = REG_SZ
                    .eOperation = REG_ADDMOD
                End With
            End With
            With .atRegKeys(1)
                .sKey = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\" & Application.Version & "\Excel\EnableedShortcutKeysCheckBoxes"
                ReDim Preserve .atRegValues(1)
                With .atRegValues(0)
                    .sName = "F1key"
                    .sData = 1
                    .eType = REG_DWORD
                    .eOperation = REG_ADDMOD
                End With
                With .atRegValues(1)
                    .sName = "F2"
                    .sData = "aa,bb"
                    .eType = REG_BINARY
                    .eOperation = REG_ADDMOD
                End With
            End With
            With .atRegKeys(2)
                .sKey = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\" & Application.Version & "\Excel\DisabledShortcutKeysCheckBoxes"
                .eOperation = REG_DELETE
            End With
            With .atRegKeys(3)
                .sKey = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\" & Application.Version & "\Excel\EnableedShortcutKeysCheckBoxes"
                ReDim Preserve .atRegValues(0)
                With .atRegValues(0)
                    .sName = "F1key"
                    .eOperation = REG_DELETE
                End With
            End With
        End With
        
        Call DisableRegWrite
        Call SetRegistry(sRegFilePath, tRegStruct)
    End Sub

'********************************************************************************
'* 内部関数定義
'********************************************************************************
' ==================================================================
' = 概要    レジストリ 値取得
' = 引数    atRegKeys       [in]    T_REG_KEYS()    キー
' = 戻値                    [out]   String          値
' = 覚書    なし
' = 依存    Mng_Registory.bas/GetRegValuesText()
' = 所属    Mng_Registory.bas
' ==================================================================
Private Function GetRegKeysText( _
    ByRef atRegKeys() As T_REG_KEYS _
) As String
    Dim sOutText As String
    sOutText = ""
    If Sgn(atRegKeys) = 0 Then
        'Do Nothing
    Else
        Dim lIdx As Long
        For lIdx = 0 To UBound(atRegKeys)
            If atRegKeys(lIdx).eOperation = REG_ADDMOD Then
                sOutText = sOutText & "[" & atRegKeys(lIdx).sKey & "]" & vbCrLf
                sOutText = sOutText & GetRegValuesText(atRegKeys(lIdx).atRegValues)
            Else
                sOutText = sOutText & "[-" & atRegKeys(lIdx).sKey & "]" & vbCrLf
            End If
            sOutText = sOutText & vbCrLf
        Next lIdx
    End If
    GetRegKeysText = sOutText
End Function

' ==================================================================
' = 概要    レジストリ キー取得
' = 引数    atRegValues     [in]    T_REG_VALUES()  値
' = 戻値                    [out]   String          キー
' = 覚書    なし
' = 依存    なし
' = 所属    Mng_Registory.bas
' ==================================================================
Private Function GetRegValuesText( _
    ByRef atRegValues() As T_REG_VALUES _
) As String
    Dim sOutText As String
    Dim sTmpText As String
    sOutText = ""
    If Sgn(atRegValues) = 0 Then
        'Do Nothing
    Else
        Dim lIdx As Long
        For lIdx = 0 To UBound(atRegValues)
            With atRegValues(lIdx)
                If .eOperation = REG_ADDMOD Then
                    Select Case .eType
                        Case REG_SZ:
                            sTmpText = """" & .sName & """=""" & .sData & """"
                            sOutText = sOutText & Replace(sTmpText, "\", "\\")
                        Case REG_DWORD:
                            sOutText = sOutText & """" & .sName & """=dword:" & .sData
                        Case REG_BINARY:
                            sOutText = sOutText & """" & .sName & """=hex:" & .sData
                        Case Else:
                            Debug.Assert 0
                    End Select
                Else
                    sOutText = sOutText & """" & .sName & """=-"
                End If
            End With
            sOutText = sOutText & vbCrLf
        Next lIdx
    End If
    GetRegValuesText = sOutText
End Function

