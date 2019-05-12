Option Explicit

'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\_lib\Excel.vbs" )
Call Include( sMyDirPath & "\_lib\String.vbs" )

'==========================================================
'= 本処理
'==========================================================
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim sFilePath
If WScript.Arguments.Count = 0 Then
	sFilePath = objWshShell.SpecialFolders("Desktop") & "\temp.xlsm"
ElseIf WScript.Arguments.Count = 1 Then
	sFilePath = WScript.Arguments(0)
	Dim sFileExt
	sFileExt = GetFileExt( sFilePath )
	Select Case sFileExt
		Case "xlsx":    'Do Nothing
		Case "xls":     'Do Nothing
		Case "xlsm":    'Do Nothing
		Case Else:
			MsgBox "Excelファイルではありません！"
			MsgBox "処理を中断します"
			WScript.Quit
	End Select
	sFilePath = WScript.Arguments(0)
Else
	MsgBox "２つ以上の引数は指定できません"
	MsgBox "処理を中断します"
	WScript.Quit
End If
Call CreateNewExcelFile( sFilePath )

WScript.CreateObject("WScript.Shell").Run sFilePath, 1, True

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
