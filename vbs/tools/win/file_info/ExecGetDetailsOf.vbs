'TODO:要コマンド化
'現状は、GetDetailsOfGetDetailsOf()のテスト用スクリプトの位置づけ。
'ファイルパスを引数で受け取るようにすれば、汎用化可能。
'合わせてコマンド化したい。

Option Explicit
'==========================================================
'= インクルード
'==========================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )  'GetDetailsOfGetDetailsOf()

'==========================================================
'= 本処理
'==========================================================
'GetDetailsOf()の詳細情報（要素番号、タイトル情報、型名、データ）を取得する
Dim sTrgtFilePath
sTrgtFilePath = "Z:\300_Musics\200_Reggae@Jamaica\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3"

Dim sLogFilePath
sLogFilePath = WScript.CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\track_title_names.txt"

Call GetDetailsOfGetDetailsOf( sTrgtFilePath, sLogFilePath )

WScript.CreateObject("WScript.Shell").Run sLogFilePath, 1, True

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

