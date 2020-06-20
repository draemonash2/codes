Option Explicit

'==========================================================
'= インクルード
'==========================================================
Call Include( "C:\codes\vbs\_lib\Windows.vbs" )     'RunasCheck()

'==========================================================
'= 設定値
'==========================================================
'Const TRGT_NETWORKDRIVE_PATH = "\\RASPBERRYPI\pockethdd"
Const TRGT_NETWORKDRIVE_PATH = "\\RASPBERRYPI\LogitecHdd3T"
Const SEARCH_VOLUME_LAVEL = "PocketHdd"
Const BACKUP_PATH_SRC = "C:\Users\draem_000\AppData\Roaming\Apple Computer\MobileSync"
Const BACKUP_PATH_DST = "700_Evacuate_iTunes\MobileSync"

'==========================================================
'= 本処理
'==========================================================
'本スクリプトを管理者として実行させる
Call RunasCheck

Dim sTrgtDrvPath
sTrgtDrvPath = ""
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'**** ネットワークドライブを探す ****
On Error Resume Next
Dim lDrvLttrIdx
Dim sDrvLttr
Dim lDrvLttrAscStrt
Dim lDrvLttrAscLast
lDrvLttrAscStrt = asc("A")
lDrvLttrAscLast = asc("Z")
For lDrvLttrIdx = lDrvLttrAscStrt to lDrvLttrAscLast
    sDrvLttr = Chr(lDrvLttrIdx)
    If Err.Number = 0 Then
        If objFSO.DriveExists(sDrvLttr) Then
            Dim objDrive
            Set objDrive = objFSO.GetDrive(sDrvLttr)
            If objDrive.IsReady = True Then
                If LCase( objDrive.VolumeName ) = LCase( SEARCH_VOLUME_LAVEL ) Then
                    sTrgtDrvPath = objDrive.Path
                    Exit For
                Else
                    'Do Nothing
                End If
            Else
                'Do Nothing
            End If
        Else
            'Do Nothing
        End If
    Else
        MsgBox "{error] " & Err.Description
        WScript.Quit
    End If
Next
On Error Goto 0

'**** ローカルドライブを探す ****
If sTrgtDrvPath = "" Then
    If objFSO.FolderExists( TRGT_NETWORKDRIVE_PATH ) Then
        sTrgtDrvPath = TRGT_NETWORKDRIVE_PATH
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'**** シンボリックリンク作成 ****
Dim oWsh
Set oWsh = WScript.CreateObject("WScript.Shell")
If sTrgtDrvPath = "" Then
    MsgBox "対象フォルダが見つかりませんでした。"
Else
    'シンボリックリンク削除
    If objFSO.FolderExists( BACKUP_PATH_SRC ) Then
        oWsh.Run "%ComSpec% /c rmdir """ & BACKUP_PATH_SRC & """"
    Else
        'Do Nothing
    End If
    
    'シンボリックリンク削除
    oWsh.Run "%ComSpec% /c mklink /d """ & BACKUP_PATH_SRC & """ """ & _
             sTrgtDrvPath & "\" & BACKUP_PATH_DST & """"
    MsgBox "iTunes バックアップ格納先を以下に設定しました。" & vbNewLine & _
           "  格納先：" & Replace( sTrgtDrvPath, ":", "" )
End If

'==========================================================
'= インクルード関数
'==========================================================
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

