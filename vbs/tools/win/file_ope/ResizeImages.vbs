'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )          ' ShowFolderSelectDialog()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )              ' ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\SettingFileClass.vbs" )    ' SettingFile

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sPROG_NAME = "画像リサイズ"
Const sRESIZE_RATE = "80%"
Const sRESIZE_PRG_ENV_NAME = "MYEXEPATH_IMAGEMAGICK"
Const sBACKUP_DIR_NAME = "_resize_img_bak"

'===============================================================================
'= 本処理
'===============================================================================
If bEXEC_TEST = True Then '{{{
    Call Test_Main()
Else
    Call Main()
End If '}}}

'===============================================================================
'= メイン関数
'===============================================================================
Public Sub Main()
    Dim sTrgtFilePath
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    
    '*** 選択ファイル取得 ***
    If EXECUTION_MODE = 0 Then 'Explorerから実行
        Dim sArg
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        For Each sArg In WScript.Arguments
            cSelectedPaths.add sArg
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finderから実行
        sSrcParDirPath = WScript.Env("Current")
        Set cSelectedPaths = WScript.Col( WScript.Env("Selected") )
    Else 'デバッグ実行
        MsgBox "デバッグモードです。"
        sSrcParDirPath = "C:\Users\draem\OneDrive\デスクトップ"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "C:\Users\draem\OneDrive\デスクトップ\parts_group_01.jpg"
        cSelectedPaths.Add "C:\Users\draem\OneDrive\デスクトップ\parts_group_02.jpg"
    End If
    
    Dim sResizePrgPath
    sResizePrgPath = objWshShell.ExpandEnvironmentStrings("%" & sRESIZE_PRG_ENV_NAME & "%")
    
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        If objFSO.FileExists(sSelectedPath) = False Then
            MsgBox "ファイルが見つかりません: " & sSelectedPath, vbExclamation, sPROG_NAME
        Else
            '--- バックアップパス作成（存在していれば "_" を追加し続ける） ---
            Dim sDir, sBase, sExt
            Dim sBackupBase, sBackupPath
            
            sDir  = objFSO.GetParentFolderName(sSelectedPath) & "\" & sBACKUP_DIR_NAME
            sBase = objFSO.GetBaseName(sSelectedPath)
            sExt  = "." & objFSO.GetExtensionName(sSelectedPath)
            
            If Not objFSO.FolderExists(sDir) Then
                objFSO.CreateFolder(sDir)
            End If
            
            sBackupBase = sBase & "_"
            sBackupPath = objFSO.BuildPath(sDir, sBackupBase & sExt)
            Do While objFSO.FileExists(sBackupPath)
                sBackupBase = sBackupBase & "_"
                sBackupPath = objFSO.BuildPath(sDir, sBackupBase & sExt)
            Loop
            
            '--- バックアップ作成 ---
            objFSO.CopyFile sSelectedPath, sBackupPath, False
            
            '--- ImageMagickで縮小（入力=バックアップ / 出力=元パスで上書き） ---
            Dim sCmd, objResult
            Set objWshShell = CreateObject("WScript.Shell")
            
            sCmd = """" & sResizePrgPath & """" _
                 & " """ & sSelectedPath & """" _
                 & " -resize """ & sRESIZE_RATE & """" _
                 & " """ & sSelectedPath & """"
            
            objResult = objWshShell.Run(sCmd, 0, True)
            If objResult <> 0 Then
                MsgBox "縮小に失敗しました。(result=" & objResult & ")" & vbCrLf _
                     & "cmd: " & sCmd, vbExclamation, sPROG_NAME
            End If
        End If
    Next
    
End Sub

'===============================================================================
'= テスト関数
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTestCase = 1
    MsgBox "=== test start ==="
    Select Case lTestCase
        Case 1
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
