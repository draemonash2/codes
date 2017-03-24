'分割されたバイナリファイルを結合する。
'結合したいバイナリファイルを選択して
'ドラッグアンドドロップすることで、
'このスクリプトと同じフォルダに「JoinFile」を作成する。
'zip ファイルの場合、ファイル名を「JoinFile」⇒ 「JoinFile.zip」
'に変更してから解凍すること。

Option Explicit

Dim asWSArgs
Set asWSArgs = WScript.Arguments

Dim sFileStr
sFileStr = ""
If asWSArgs.Count <= 1 Then
    WScript.Echo "More Arguments!"
    WScript.Quit
Else
    sFileStr = asWSArgs( 0 )
    Dim lArgIdx
    For lArgIdx = 1 to ( asWSArgs.Count - 1 )
        sFileStr = sFileStr & " + " & asWSArgs( lArgIdx )
    Next
End If

Dim oFileSys
Dim sParentDirName
Dim sOutFileName
Set oFileSys = CreateObject("Scripting.FileSystemObject")
sParentDirName = oFileSys.GetParentFolderName(asWSArgs(0))
sOutFileName = "JoinFile"

Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c copy /b " & sFileStr & " " & sParentDirName & "\" & sOutFileName, 0, false
Set objShell = Nothing

