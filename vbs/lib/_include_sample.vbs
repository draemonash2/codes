'==========================================================
'= インクルード
'==========================================================
Dim objWshShell
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
Call Include( objWshShell.CurrentDirectory & "\String.vbs" )

'==========================================================
'= 本処理
'==========================================================
'★ここに処理を書く★

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
