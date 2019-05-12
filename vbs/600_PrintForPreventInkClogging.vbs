'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\_lib\Excel.vbs" )

'==========================================================
'= 本処理
'==========================================================
Call PrintExcelSheet( _
    "C:\Users\draem_000\Documents\Dropbox\100_Documents\111_【生活】＜電化製品＞PC＆携帯電話\PC\プリンタインク詰まり防止用印刷ページ.xlsx", _
    "Sheet1", _
    1, _
    1, _
    1 _
)

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
