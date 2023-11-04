Option Explicit

'TODO:要コマンド化
'TODO:EXE_PATH_ORG/EXE_PATH_NEW 引数化

' 概要）
'指定されたフォルダ配下に存在するショートカット(*.lnk)の指示先を、置換する。
'
'実行方法）
'1. 設定値を修正する。
'2. 本スクリプトファイルを実行する。

'==========================================================
'= インクルード
'==========================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )  'GetFileList2()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Log.vbs" )         'class LogMng

'==========================================================
'= 設定値
'==========================================================
Const EXE_PATH_ORG = "C:\Users\draem\Programs\program\prg_exe"
Const EXE_PATH_NEW = "C:\prg_exe"

'==========================================================
'= 本処理
'==========================================================
MsgBox "本プログラムは管理者権限が必要となる場合があります。" & vbNewLine & "エラーが発生した場合、管理者権限にて実行してください。"

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim sTrgtDir
sTrgtDir = objWshShell.CurrentDirectory
'sTrgtDir = objWshShell.SpecialFolders("StartMenu")
Dim bIsContinue
Do
    Dim vAnswer
    vAnswer = MsgBox( "以下を対象に実行します。実行しますか？" & vbNewLine & sTrgtDir, vbOkCancel )
    If vAnswer = vbOk Then
        bIsContinue = False
    Else
        vAnswer = MsgBox( "処理を続けますか？", vbOkCancel )
        If vAnswer = vbCancel Then
            WScript.Quit
        Else
            sTrgtDir = InputBox ( "対象ディレクトリを指定してください。" )
            bIsContinue = True
        End If
    End If
Loop While bIsContinue = True

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sMyFileBaseName
sMyFileBaseName = objFSO.GetBaseName( WScript.ScriptFullName )

Dim oLogMng
Set oLogMng = New LogMng
Dim sLogFilePath
sLogFilePath = sTrgtDir & "\" & sMyFileBaseName & ".log"
Call oLogMng.Open( sLogFilePath, "w" )

Dim asFileList
Call GetFileList2( sTrgtDir, asFileList, 1 )

oLogMng.Puts( "target directory path : " & sTrgtDir )
oLogMng.Puts( "org path              : " & EXE_PATH_ORG )
oLogMng.Puts( "new path              : " & EXE_PATH_NEW )
oLogMng.Puts( "" )
oLogMng.Puts( "### Legend ###" )
oLogMng.Puts( "  Replaced : replaced target pathes" )
oLogMng.Puts( "  UnMatch  : replace word is nothing at target path" )
oLogMng.Puts( "  NoShrtCt : not a program shortcut file" )
oLogMng.Puts( "### Result ###" )
oLogMng.Puts( "<Result>" & chr(9) & "<sFileDirPath>" & chr(9) & "<sOrgDirPath>" & chr(9) & "<sNewDirPath>" )

Dim i
For i = 0 to UBound( asFileList ) - 1
    Dim sFileDirPath
    sFileDirPath = asFileList(i)
    
    Dim sOrgDirPath
    Dim sNewDirPath
    If objFSO.GetExtensionName( sFileDirPath ) = "lnk" Then
        With objWshShell.CreateShortcut( sFileDirPath )
            sOrgDirPath = .TargetPath
            If InStr( sOrgDirPath, EXE_PATH_ORG ) > 0 Then
                sNewDirPath = Replace( sOrgDirPath, EXE_PATH_ORG, EXE_PATH_NEW )
                    .TargetPath = sNewDirPath
                    .Save
                oLogMng.Puts( "[Replaced]" & chr(9) & sFileDirPath & chr(9) & sOrgDirPath & chr(9) & sNewDirPath )
            Else
                oLogMng.Puts( "[UnMatch ]" & chr(9) & sFileDirPath & chr(9) & sOrgDirPath )
            End If
        End With
    Else
        oLogMng.Puts( "[NoShrtCt]" & chr(9) & sFileDirPath )
    End If
Next

oLogMng.Close()
Set oLogMng = Nothing

MsgBox _
    "以下にプログラムショートカットの指示先置換結果を出力しました。" & vbNewLine & _
    "  " & sLogFilePath

'==========================================================
'= インクルード関数
'==========================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

