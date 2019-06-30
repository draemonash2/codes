Option Explicit

' ReplaceStrInTxtFile.vbs <search_word> <replace_word> <target_file_path>

'===============================================================================
'= インクルード
'===============================================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\Collection.vbs" )

'===============================================================================
'= 本処理
'===============================================================================

Dim sSearchStr
Dim sReplaceStr
Dim sTrgtFilePath
Dim bIsRegExp
if Wscript.Arguments.Count = 3 then
	sSearchStr		= Wscript.Arguments(0)
	sReplaceStr		= Wscript.Arguments(1)
	sTrgtFilePath	= Wscript.Arguments(2)
else
	wscript.echo "arguments error!"
	wscript.quit
end if

Dim cInputFile
Set cInputFile = CreateObject("System.Collections.ArrayList")
Dim cOutputFile
Set cOutputFile = CreateObject("System.Collections.ArrayList")

call ReadTxtFileToCollection(sTrgtFilePath, cInputFile)

Dim bIsMatch
bIsMatch = false
Dim sLine
for each sLine in cInputFile
	if instr(sLine, sSearchStr) > 0 then
		bIsMatch = true
		cOutputFile.add replace(sLine, sSearchStr, sReplaceStr)
	else
		cOutputFile.add sLine
	end if
next

if bIsMatch = true then
	call WriteTxtFileFrCollection(sTrgtFilePath, cOutputFile, true)
end if

'===============================================================================
'= インクルード関数
'===============================================================================
' 外部プログラム インクルード関数
Private Function Include( _
	ByVal sOpenFile _
)
	Dim objFSO
	Dim objVbsFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	sOpenFile = objFSO.GetAbsolutePathName( sOpenFile )
	Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
	
	ExecuteGlobal objVbsFile.ReadAll()
	objVbsFile.Close
	
	Set objVbsFile = Nothing
	Set objFSO = Nothing
End Function
