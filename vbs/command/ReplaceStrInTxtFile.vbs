Option Explicit

'<usage>
'  ReplaceStrInTxtFile.vbs <search_word> <replace_word> <target_file_path>

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Collection.vbs" ) 'ReadTxtFileToCollection()
                                                            'WriteTxtFileFrCollection()

'===============================================================================
'= 本処理
'===============================================================================

Dim sSearchStr
Dim sReplaceStr
Dim sTrgtFilePath
Dim bIsRegExp
if Wscript.Arguments.Count = 3 then
    sSearchStr      = Wscript.Arguments(0)
    sReplaceStr     = Wscript.Arguments(1)
    sTrgtFilePath   = Wscript.Arguments(2)
else
    WScript.Echo "[error] argment num = " & WScript.Arguments.Count
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
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

