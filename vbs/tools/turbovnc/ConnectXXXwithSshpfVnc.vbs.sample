Option Explicit

'<<概要>>
'  SSHポートフォワードをバックグラウンドで実行し、TurboVNC-Viewerで接続する
'  
'<<使用方法>>
'  本プログラムを実行する
'  
'<<仕様>>
'  ・特になし

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Windows.vbs" )    'ExistProcess()

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sSCRIPT_NAME = "TODO:接続＠TurboVNC"

Const sVNC_VIEWER_PATH = "C:\prg\TurboVNC\vncviewer.bat"
Const sCONNECT_NAME = "TODO:"
Const sUSER_NAME = "TODO:"
Const sSRC_PORT_NO = "TODO:"
Const sDST_PORT_NO = "TODO:"

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
    Dim sCmd
    If Not ExistProcess("ssh.exe") Then
        sCmd = "ssh " & sUSER_NAME & "@" & sCONNECT_NAME & " -L "  & sSRC_PORT_NO & ":localhost:" & sDST_PORT_NO & " -N"
        CreateObject("WScript.Shell").Run sCmd, 0
    End If
    sCmd = sVNC_VIEWER_PATH & " localhost::" & sSRC_PORT_NO
    CreateObject("WScript.Shell").Run sCmd, 0
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
