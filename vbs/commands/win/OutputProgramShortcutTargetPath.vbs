Option Explicit

' 概要）
'指定されたフォルダ配下に存在するショートカット(*.lnk)の指示先を、
'一覧化して出力する。
'
'実行方法）
'1. 本スクリプトファイルを実行する。

'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\_lib\FileSystem.vbs" )
Call Include( sMyDirPath & "\_lib\Log.vbs" )

'==========================================================
'= 設定値
'==========================================================

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
oLogMng.Puts( "" )
oLogMng.Puts( "### Result ###" )
oLogMng.Puts( "<Type>" & chr(9) & "<sFileDirPath>" & chr(9) & "<sTargetPath>" )

Dim i
For i = 0 to UBound( asFileList ) - 1
    Dim sFileDirPath
    sFileDirPath = asFileList(i)
    
    If objFSO.GetExtensionName( sFileDirPath ) = "lnk" Then
        With objWshShell.CreateShortcut( sFileDirPath )
            oLogMng.Puts( "[ShrtCt  ]" & chr(9) & sFileDirPath & chr(9) & .TargetPath )
        End With
    Else
        oLogMng.Puts( "[NoShrtCt]" & chr(9) & sFileDirPath )
    End If
Next

oLogMng.Close()
Set oLogMng = Nothing

MsgBox _
    "以下にプログラムショートカットの指示先を出力しました。" & vbNewLine & _
    "  " & sLogFilePath

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

