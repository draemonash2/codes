Option Explicit
'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\_lib\Log.vbs" )
Call Include( sMyDirPath & "\_lib\Windows.vbs" )

'==========================================================
'= 本処理
'==========================================================
On Error Resume Next

'本スクリプトを管理者として実行させる
If ExecRunas( False ) Then WScript.Quit

Const ENV_NAME = "Path"         'ロジック上「Path」以外は指定できないため、変更不可
Const TARGET_GROUP = "System"   'System：すべてのユーザー 、User：現在のユーザー 、Volatile：現在のログオン 、Process：現在のプロセス

If WScript.Arguments.Count > 1 Then
    Dim sEnvValue
    sEnvValue = WScript.Arguments(1) 'WScript.Arguments(0)は、ExecRunas()にて使用される
Else
    WScript.Echo "引数が指定されていません"
    WScript.Echo "処理を中断します"
    WScript.Quit
End If
'MsgBox sEnvValue

'Dim oLog
'Set oLog = New LogMng
'oLog.Open "C:\Users\draem_000\Desktop\test.log", "w"

Dim objUsrEnv
If Err.Number = 0 Then
    Set objUsrEnv = WScript.CreateObject("WScript.Shell").Environment(TARGET_GROUP)
    If Err.Number = 0 Then
        Dim sEnvValOrg
        Dim sEnvValNew
        sEnvValOrg = objUsrEnv.Item(ENV_NAME)
        If InStr( sEnvValOrg, sEnvValue ) > 0 Then
            WScript.Echo ENV_NAME & "に存在します"
        Else
            sEnvValNew = sEnvValOrg & ";" & sEnvValue
        '   oLog.Puts sEnvValOrg
        '   oLog.Puts sEnvValNew
            objUsrEnv.Item(ENV_NAME) = sEnvValNew
            WScript.Echo ENV_NAME & "に追加しました"
        End If
    Else
        WScript.Echo "エラー: " & Err.Description
    End If
Else
    WScript.Echo "エラー: " & Err.Description
End If

Set objUsrEnv = Nothing

'oLog.Close
'Set oLog = Nothing

'==========================================================
'= 関数定義
'==========================================================
' 外部プログラム インクルード関数
Function Include( _
    ByVal sOpenFile _
)
    Dim objFSO
    Dim objVbsFile
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function
