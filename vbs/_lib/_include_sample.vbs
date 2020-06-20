'===============================================================================
'= インクルード
'===============================================================================
Call Include( "C:\codes\vbs\_lib\String.vbs" ) '★()

'===============================================================================
'= 本処理
'===============================================================================
'★ここに処理を書く★

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( _
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
