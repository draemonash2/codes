Option Explicit

'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\lib\FileSystem.vbs" )
Call Include( sMyDirPath & "\lib\Log.vbs" )

'==========================================================
'= 設定値
'==========================================================
Const EXE_PATH_ORG = "C:\Users\draem_000\Documents\Amazon Drive\100_Programs\program\prg_exe"
Const EXE_PATH_NEW = "C:\prg_exe"

'==========================================================
'= 本処理
'==========================================================
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sTrgtDir
Dim sMyFileBaseName
sTrgtDir = objWshShell.SpecialFolders("StartMenu")
sMyFileBaseName = objFSO.GetBaseName( WScript.ScriptFullName )

Dim oLogMng
Set oLogMng = New LogMng
Call oLogMng.LogFileOpen( sTrgtDir & "\" & sMyFileBaseName & ".log", "w" )

Dim asFileList
Call GetFileList2( sTrgtDir, asFileList, 1 )

oLogMng.LogPuts( "[Result]" & chr(9) & "[sFileDirPath]" & chr(9) & "[sOrgDirPath]" & chr(9) & "[sNewDirPath]" )

Dim i
For i = 0 to UBound( asFileList ) - 1
	Dim sFileDirPath
	sFileDirPath = asFileList(i)
	
	Dim sOrgDirPath
	Dim sNewDirPath
	If objFSO.GetExtensionName( sFileDirPath ) = "lnk" Then
		sOrgDirPath = objWshShell.CreateShortcut( sFileDirPath ).TargetPath
		sNewDirPath = Replace( sOrgDirPath, EXE_PATH_ORG, EXE_PATH_NEW )
		objWshShell.CreateShortcut( sFileDirPath ).TargetPath = sNewDirPath
		oLogMng.LogPuts( "[Replaced]" & chr(9) & sFileDirPath & chr(9) & sOrgDirPath & chr(9) & sNewDirPath )
	Else
		oLogMng.LogPuts( "[Stay    ]" & chr(9) & sFileDirPath )
	End If
Next

oLogMng.LogFileClose()
Set oLogMng = Nothing

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

