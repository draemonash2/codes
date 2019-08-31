Option Explicit
'参考URL:http://d.hatena.ne.jp/maeyan/touch/20140416/1397602663

' progrress bar cscript class v1.00

' = 依存    なし
' = 所属    ProgressBarCscript.vbs
Class ProgressBar
    Private sStatus
    Private objFSO
    Private objWshShell
    
    Private Sub Class_Initialize
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objWshShell = WScript.CreateObject("WScript.Shell")
        Dim sExeFileName
        sExeFileName = LCase(objFSO.GetFileName(WScript.FullName))
        if sExeFileName = "cscript.exe" then
            'Do Nothing
        else
            objWshShell.Run "cscript //nologo """ & Wscript.ScriptFullName & """", 1, False
            Wscript.Quit
        end if
    End Sub
    
    Private Sub Class_Terminate
        Set objFSO = Nothing
        Set objWshShell = Nothing
    End Sub
    
    ' ==================================================================
    ' = 概要    メッセージを更新する
    ' = 引数    sProgMsg      String   [in] メッセージ
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Property Let Message( _
        ByVal sMessage _
    )
        if sStatus = "Update" then
            Wscript.StdOut.Write vbCrLf
        end if
        Wscript.StdOut.Write sMessage & vbCrLf
        sStatus = "Message"
    End Property
    
    ' ==================================================================
    ' = 概要    進捗を更新する
    ' = 引数    lBunsi      Long   [in] 進捗
    ' = 引数    lBunbo      Long   [in] 進捗最大値
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Sub Update( _
        ByVal lBunsi, _
        ByVal lBunbo _
    )
        'パーセンテージ計算
        Dim iPercentage
        Dim sPercentage
        iPercentage = Cint((lBunsi / lBunbo) * 100)
        sPercentage = iPercentage & "%"
        sPercentage = String(4 - Len(sPercentage), " ") & sPercentage
        
        '進捗バー
        Dim sProgressBar
        sProgressBar = String(Cint(iPercentage/5), "=") & ">" & String(20 - Cint(iPercentage/5), " ")
        
        '描画
        Wscript.StdOut.Write sPercentage & " |" & sProgressBar & "| " & lBunsi & "/" & lBunbo & vbCr
        sStatus = "Update"
    End Sub
    
    ' ==================================================================
    ' = 概要    プログレスバーを終了する
    ' = 引数    なし
    ' = 戻値    なし
    ' = 覚書    cscriptは終了できない
    ' ==================================================================
'   Public Function Quit()
'       gobjExplorer.Document.Body.Style.Cursor = "default"
'       gobjExplorer.Quit
'   End Function
    
End Class
    If WScript.ScriptName = "ProgressBarCscript.vbs" Then
        Call Test_ProgressBar
    End If
    Private Sub Test_ProgressBar
        Dim lProcIdx
        Dim lProcNum
        Dim objPrgrsBar
        Set objPrgrsBar = New ProgressBar
        
        '#処理１
        objPrgrsBar.Message = "長い処理 実行!"
        lProcNum = 255
        For lProcIdx = 1 To lProcNum
            WScript.Sleep 1
            Call objPrgrsBar.Update(lProcIdx, lProcNum)
        Next
        
        '#処理２
        objPrgrsBar.Message = "短い処理 実行!"
        lProcNum= 10
        For lProcIdx = 1 To lProcNum
            WScript.Sleep 45
            Call objPrgrsBar.Update(lProcIdx, lProcNum)
        Next
        
        objPrgrsBar.Message = "Complete!!"
        msgbox "終了しました"
    End Sub
