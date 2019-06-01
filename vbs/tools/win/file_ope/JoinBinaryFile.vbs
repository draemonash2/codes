'分割されたバイナリファイルを結合する。
'結合したいバイナリファイルを選択して
'ドラッグアンドドロップすることで、
'このスクリプトと同じフォルダに「JoinFile」を作成する。
'zip ファイルの場合、ファイル名を「JoinFile」⇒ 「JoinFile.zip」
'に変更してから解凍すること。
'
'usage
' cscript.exe .\JoinBinaryFile.vbs <file_path> <file_path> [<file_path>]...

Option Explicit

'==========================================================
'= 本処理
'==========================================================
Dim sFileStr
sFileStr = ""
If WScript.Arguments.Count <= 1 Then
    WScript.StdOut.WriteLine "More Arguments!"
    WScript.Quit
Else
    sFileStr = WScript.Arguments( 0 )
    Dim lArgIdx
    For lArgIdx = 1 to ( WScript.Arguments.Count - 1 )
        sFileStr = sFileStr & " + " & WScript.Arguments( lArgIdx )
    Next
End If

Dim oFileSys
Dim sParentDirName
Dim sOutFileName
Set oFileSys = CreateObject("Scripting.FileSystemObject")
sParentDirName = oFileSys.GetParentFolderName(WScript.Arguments(0))
sOutFileName = "JoinFile"

Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c copy /b " & sFileStr & " " & sParentDirName & "\" & sOutFileName, 0, false
Set objShell = Nothing
