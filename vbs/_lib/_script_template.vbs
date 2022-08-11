Option Explicit

'<<概要>>
'  
'  
'<<使用方法>>
'  
'  
'<<仕様>>
'  ・

'===============================================================================
'= インクルード
'===============================================================================
'Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sSCRIPT_NAME = ★

'===============================================================================
'= 本処理
'===============================================================================
Dim cArgs '{{{
Set cArgs = CreateObject("System.Collections.ArrayList")

If bEXEC_TEST = True Then
    Call Test_Main()
Else
    Dim vArg
    For Each vArg in WScript.Arguments
        cArgs.Add vArg
    Next
    Call Main()
End If '}}}

'===============================================================================
'= メイン関数
'===============================================================================
Public Sub Main()
    Dim sTrgtPath
    Dim lBakFileNumMax
    Dim sBakLogFilePath
    If cArgs.Count >= 1 Then
        sTrgtPath = cArgs(0)
    Else
        WScript.Echo "引数を指定してください。プログラムを中断します。"
        Exit Sub
    End If
End Sub

'===============================================================================
'= 内部関数
'===============================================================================

'===============================================================================
'= テスト関数
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTestCase = 1
    
    cArgs.Add sTrgtFilePath
    
    MsgBox "=== test start ==="
    
    Select Case lTestCase
        Case 1
            Call Main()
            MsgBox "1 実行後"
        Case Else
            Call Main()
    End Select
    
    MsgBox "=== test finished ==="
End Sub '}}}

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile ) '{{{
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function '}}}

