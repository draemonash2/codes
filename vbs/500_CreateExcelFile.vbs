Option Explicit

'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\lib\Excel.vbs" )
Call Include( sMyDirPath & "\lib\FileSystem.vbs" )

'==========================================================
'= �{����
'==========================================================
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim sFilePath
If WScript.Arguments.Count = 0 Then
	sFilePath = objWshShell.SpecialFolders("Desktop") & "\temp.xlsm"
Else
	Dim sFileExt
	sFileExt = GetFileExt( sFilePath )
	Select Case sFileExt
		Case "xlsx":    'Do Nothing
		Case "xls":     'Do Nothing
		Case "xlsm":    'Do Nothing
		Case Else:
			MsgBox "Excel�t�@�C���ł͂���܂���I"
			MsgBox "�����𒆒f���܂�"
			WScript.Quit
	End Select
	sFilePath = WScript.Arguments(0)
End If
Call CreateNewExcelFile( sFilePath )

WScript.CreateObject("WScript.Shell").Run sFilePath, 1, True

'==========================================================
'= �֐���`
'==========================================================
' �O���v���O���� �C���N���[�h�֐�
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