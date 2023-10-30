Option Explicit

'<<概要>>
'  TODO:
'  
'<<使用方法>>
'  TODO:
'  
'<<仕様>>
'  ・TODO:

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" )    'ExistProcess()

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sSCRIPT_NAME = "WSL2接続＠TeraTerm"

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
    If Not ExistProcess("wsl.exe") Then
        CreateObject("Wscript.Shell").Run "cmd /c wsl", 0
    End If
    CreateObject("Wscript.Shell").Run "cmd /c C:\codes\ttl\login_wsl2.ttl", 0
End Sub

'===============================================================================
'= 内部関数
'===============================================================================

'===============================================================================
'= テスト関数
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTestCase = 1
    
    MsgBox "=== test start ==="
    
    'Select Case lTestCase
    '    Case 1
    '        Call Main()
    '        MsgBox "1 実行後"
    '    Case Else
    '        Call Main()
    'End Select
    
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

