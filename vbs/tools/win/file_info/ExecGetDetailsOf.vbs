'TODO:要コマンド化

Option Explicit
'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )

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
