Attribute VB_Name = "Mng_SpecialPaste"
Option Explicit

' specialpaste manage library v1.01

Const EXEC_SEND_KEY As Boolean = False
Const SENDKEY_SLEEPTIME As Long = 50

' ==================================================================
' = 概要    条件付き書式対応モード切り替え
' = 引数    なし
' = 戻値    なし
' = 依存    Mng_SpecialPaste.bas/NewInsert()
' =         Mng_SpecialPaste.bas/NewPaste()
' =         Mng_SpecialPaste.bas/NewDelete()
' = 所属    Mng_SpecialPaste.bas
' ==================================================================
Public Sub EnableSpetialKeyMode()
    MsgBox "以下のショートカットキーを「条件付き書式対応モード」に切り替えます。" & vbNewLine & _
           "・Shift + Ctrl + ""+""" & vbNewLine & _
           "・Ctrl + v" & vbNewLine & _
           "・Ctrl + ""-""" & vbNewLine & _
           "" & vbNewLine & _
           "★注意★ このモードではアンドゥできません！"
    Application.OnKey "+^{+}", "NewInsert"
    Application.OnKey "^v", "NewPaste"
    Application.OnKey "^-", "NewDelete"
End Sub

' ==================================================================
' = 概要    ノーマルモード切り替え
' = 引数    なし
' = 戻値    なし
' = 依存    なし
' = 所属    Mng_SpecialPaste.bas
' ==================================================================
Public Sub DisableSpetialKeyMode()
    MsgBox "以下のショートカットキーを「ノーマルモード」に切り替えます。" & vbNewLine & _
           "・Shift + Ctrl + ""+""" & vbNewLine & _
           "・Ctrl + v" & vbNewLine & _
           "・Ctrl + ""-"""
    Application.OnKey "+^{+}"
    Application.OnKey "^v"
    Application.OnKey "^-"
End Sub

' ==================================================================
' = 概要    挿入貼り付け
' = 引数    なし
' = 戻値    なし
' = 依存    Mng_SendKeys.bas/SendKeysBetweenWait()
' = 所属    Mng_SpecialPaste.bas
' ==================================================================
Private Sub NewInsert()
    Select Case Application.CutCopyMode
        Case xlCopy
            MsgBox "挿入貼り付けは無効です。"
        Case xlCut
            MsgBox "挿入貼り付けは無効です。"
        Case Else
            If EXEC_SEND_KEY = True Then
                Call SendKeysBetweenWait("%hie", SENDKEY_SLEEPTIME) '挿入
            Else
                Application.ScreenUpdating = False
                Selection.Insert
                Application.ScreenUpdating = True
            End If
    End Select
End Sub

' ==================================================================
' = 概要    挿入貼り付け
' = 引数    なし
' = 戻値    なし
' = 依存    Mng_SendKeys.bas/SendKeysBetweenWait()
' = 所属    Mng_SpecialPaste.bas
' ==================================================================
Private Sub NewPaste()
    Select Case Application.CutCopyMode
        Case xlCopy
            If EXEC_SEND_KEY = True Then
                Call SendKeysBetweenWait("%hvf", SENDKEY_SLEEPTIME) '数式貼り付け
            Else
                Application.ScreenUpdating = False
                '数式を貼り付ける
                Selection.PasteSpecial _
                    Paste:=xlPasteFormulas, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=False
    '            '条件付き書式を結合して貼り付ける
    '            Selection.PasteSpecial _
    '                Paste:=xlPasteAllMergingConditionalFormats, _
    '                Operation:=xlNone, _
    '                SkipBlanks:=False, _
    '                Transpose:=False
                Application.ScreenUpdating = True
            End If
        Case xlCut
            MsgBox "カット＆ペーストは無効です。"
        Case Else
            If EXEC_SEND_KEY = True Then
                Call SendKeysBetweenWait("%hvt", SENDKEY_SLEEPTIME) '貼り付け
            Else
                Application.ScreenUpdating = False
                Dim doDataObj As New DataObject
                doDataObj.GetFromClipboard
                Selection(1).Value = doDataObj.GetText
                Application.ScreenUpdating = True
            End If
    End Select
End Sub

' ==================================================================
' = 概要    削除
' = 引数    なし
' = 戻値    なし
' = 依存    Mng_SendKeys.bas/SendKeysBetweenWait()
' = 所属    Mng_SpecialPaste.bas
' ==================================================================
Private Sub NewDelete()
    '「一行削除」時は「行挿入→二行削除」とする。
    '（一行のみの削除は条件付き書式が増殖されてしまうため）
    If Selection.Rows.Count = 1 And _
       Selection.Columns.Count = Columns.Count Then
        MsgBox "条件付き書式が崩れるため、単行の削除はできません。"
'        If EXEC_SEND_KEY = True Then
'            Call SendKeysBetweenWait("%hie", SENDKEY_SLEEPTIME) '挿入
'            Call SendKeysBetweenWait("+{DOWN}", SENDKEY_SLEEPTIME) 'シフト+下
'            Call SendKeysBetweenWait("%hdd", SENDKEY_SLEEPTIME) '削除
'        Else
'            Application.ScreenUpdating = False
'            Selection.Insert
'            Selection.Resize(Selection.Rows.Count + 1).Select '行だけを拡張
'            Selection.Delete
'            Selection.Resize(1).Select
'            Application.ScreenUpdating = True
'        End If
    Else
        If EXEC_SEND_KEY = True Then
            Call SendKeysBetweenWait("%hdd", SENDKEY_SLEEPTIME) '削除
        Else
            Application.ScreenUpdating = False
            Selection.Delete
            Application.ScreenUpdating = True
        End If
    End If
End Sub

