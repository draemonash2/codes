Option Explicit

Private Const PROGBAR_ELAPSED_TIME_DIFF_COUNT_MAX = 10
Private Const PROGBAR_BASIC_LINE_NUM = 4
Private Const PROGBAR_WIN_WIDTH = 600
Private Const PROGBAR_REMAINING_TIME_INIT = 7200
Private Const PROGBAR_ONEDAYSEC = 86400 '60[s] * 60[m] * 24[h]
Private Const PROGBAR_ONEHOURSEC = 3600 '60[s] * 60[m]
Private Const PROGBAR_ONEMINSEC = 60    '60[s]

' = 依存	なし
' = 所属	ProgressBarIE.vbs
Class ProgressBar
    Dim gobjExplorer
    Dim glWinHeight
    Dim glWinHeightOld
    Dim gsProgMsg
    Dim gdProgPerRaw
    Dim gdProgPer10
    Dim gdProgPer100
    Dim gdStartTime
    Dim gsStartDate
    Dim gdElapsedTime
    Dim gdProgPerLastCalc
    Dim glElapsedTimeStoreNum
    Dim gdElapsedTimeDiffTable()
    Dim gdElapsedTimeLastCalc
    Dim gdRemainingTime
    
    Private Sub Class_Initialize()
        Dim objWMIService
        Dim colItems
        Dim strComputer
        Dim objItem
        Dim intHorizontal
        Dim intVertical
        strComputer = "."
        Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
        For Each objItem in colItems
            intHorizontal = objItem.ScreenWidth
            intVertical = objItem.ScreenHeight
        Next
        Set objWMIService = Nothing
        Set colItems = Nothing
        
        gsProgMsg = ""
        glWinHeight = CalcWinHeight( PROGBAR_BASIC_LINE_NUM )
        glWinHeightOld = glWinHeight
        gdProgPerRaw = 0
        gdProgPer10 = 0
        gdProgPer100 = 0
        gdStartTime = Timer()
        gsStartDate = Date()
        gdElapsedTime = 0
        gdProgPerLastCalc = 0
        glElapsedTimeStoreNum = 0
        ReDim Preserve gdElapsedTimeDiffTable(PROGBAR_ELAPSED_TIME_DIFF_COUNT_MAX - 1)
        Dim i
        For i = 0 To PROGBAR_ELAPSED_TIME_DIFF_COUNT_MAX - 1
            gdElapsedTimeDiffTable(i) = 0
        Next
        gdElapsedTimeLastCalc = 0
        gdRemainingTime = PROGBAR_REMAINING_TIME_INIT
        
        Set gobjExplorer = CreateObject("InternetExplorer.Application")
        gobjExplorer.Navigate "about:blank"
        gobjExplorer.ToolBar = 0
        gobjExplorer.StatusBar = 0
        gobjExplorer.Width = PROGBAR_WIN_WIDTH
        gobjExplorer.Height = glWinHeight
        gobjExplorer.Left = ( intHorizontal - gobjExplorer.Width ) / 2
        gobjExplorer.Top = ( intVertical - gobjExplorer.Height ) / 2
        gobjExplorer.Document.Body.InnerHTML = _
            "<font face=""ＭＳ ゴシック"">" & _
            "<span style=""font-size:18px; line-height:22px;"">" & _
            "処理中...<br>" & _
            "</span>" & _
            "</font>" & _
            ""
        gobjExplorer.Visible = 1
        
        Call ActiveIE
        gobjExplorer.Document.Body.Style.Cursor = "wait"
        gobjExplorer.Document.Title = "進捗状況"
    End Sub
    
    Private Sub Class_Terminate()
        'Do Nothing
    End Sub
    
    ' ==================================================================
    ' = 概要    タイトルを更新する
    ' = 引数    sTitle    String   [in] タイトル
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Property Let Title( _
        ByVal sTitle _
    )
        gobjExplorer.Document.Title = sTitle
    End Property
    
    ' ==================================================================
    ' = 概要    メッセージを更新する
    ' = 引数    sProgMsg      String   [in] メッセージ
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Property Let Message( _
        ByVal sProgMsg _
    )
        Dim lBrNum
        Dim lLineNum
        lBrNum = ( Len( sProgMsg ) - Len( Replace( sProgMsg, vbNewLine, "" ) ) ) / 2
        lLineNum = ( lBrNum + 1 ) + PROGBAR_BASIC_LINE_NUM + 1
        glWinHeight = CalcWinHeight( lLineNum )
        If sProgMsg = "" Then
            gsProgMsg = ""
        Else
            gsProgMsg = Replace( sProgMsg, vbNewLine, "<br>" ) & "<br><br>"
        End If
    End Property
    
    ' ==================================================================
    ' = 概要    進捗を更新する
    ' = 引数    dProgPerRaw   Double   [in] 進捗（0 ～ 1 の小数で指定）
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Function Update( _
        ByVal dProgPerRaw _
    )
        If dProgPerRaw > 1 Or dProgPerRaw < 0 Then
            MsgBox "指定されたプログレスバーの進捗が最大最小範囲外の値が指定されています！" & vbNewLine & _
                   "値：" & dProgPerRaw
            MsgBox "プログラムを中止します！"
            Call Quit
            WScript.Quit
        End If
        
        gdProgPerRaw = dProgPerRaw
        gdProgPer10 = Int( dProgPerRaw * 10 )
        gdProgPer100 = Int( dProgPerRaw * 100 )
        
        '経過時間算出
        Dim sDateOld
        Dim sDateNow
        Dim dSecondOld
        Dim dSecondNow
        Dim lDateDiff
        sDateOld = gsStartDate
        sDateNow = Date()
        dSecondOld = gdStartTime
        dSecondNow = Timer()
        lDateDiff = DateDiff("d", sDateOld, sDateNow)
        If lDateDiff > 0 Then
            gdElapsedTime = (PROGBAR_ONEDAYSEC * (lDateDiff - 1)) + (PROGBAR_ONEDAYSEC - dSecondOld) + dSecondNow
        ElseIf lDateDiff = 0 Then
            gdElapsedTime = dSecondNow - dSecondOld
        Else
            gdElapsedTime = 0
        End If
        gdElapsedTime = CDbl( gdElapsedTime )
        
        '残り時間算出
        Dim dProgPerCur
        Dim dProgPerDiff
        Dim dElapsedTimeCur
        dProgPerCur = gdProgPerRaw
        dProgPerDiff = dProgPerCur - gdProgPerLastCalc
        dElapsedTimeCur = gdElapsedTime
        If Int(dProgPerDiff * 100) >= 1 Then
            Dim dProgPerRem
            Dim dElapsedTimeDiff
            Dim dElapsedTime1PerCur
            dProgPerRem = 100 - (dProgPerCur * 100)
            dElapsedTimeDiff = dElapsedTimeCur - gdElapsedTimeLastCalc
            dElapsedTime1PerCur = dElapsedTimeDiff / Int(dProgPerDiff * 100)
            Dim dElapsedTimeSum
            dElapsedTimeSum = dElapsedTime1PerCur
            Dim i
            For i = 0 To glElapsedTimeStoreNum - 1 Step 1
                dElapsedTimeSum = dElapsedTimeSum + gdElapsedTimeDiffTable(i)
            Next
            Dim dElapsedTimeAvg
            dElapsedTimeAvg = dElapsedTimeSum / (glElapsedTimeStoreNum + 1)
            gdRemainingTime = dElapsedTimeAvg * dProgPerRem
            
            For i = PROGBAR_ELAPSED_TIME_DIFF_COUNT_MAX - 1 To 1 Step -1
                gdElapsedTimeDiffTable(i) = gdElapsedTimeDiffTable(i - 1)
            Next
            gdElapsedTimeDiffTable(0) = dElapsedTime1PerCur
            If glElapsedTimeStoreNum < PROGBAR_ELAPSED_TIME_DIFF_COUNT_MAX Then
                glElapsedTimeStoreNum = glElapsedTimeStoreNum + 1
            Else
                'Do Nothing
            End If
            gdProgPerLastCalc = dProgPerCur
            gdElapsedTimeLastCalc = dElapsedTimeCur
        ElseIf Fix(dProgPerDiff * 100) < 0 Then
            '進捗が下がったらクリアする
            For i = 0 To PROGBAR_ELAPSED_TIME_DIFF_COUNT_MAX - 1
                gdElapsedTimeDiffTable(i) = 0
            Next
            gdProgPerLastCalc = dProgPerCur
            gdElapsedTimeLastCalc = dElapsedTimeCur
            glElapsedTimeStoreNum = 0
            gdRemainingTime = PROGBAR_REMAINING_TIME_INIT
        Else
            'Do Nothing
        End If
        
        '高さ調節
        If glWinHeight = glWinHeightOld Then
            'Do Nothing
        Else
            gobjExplorer.Height = glWinHeight
        End If
        glWinHeightOld = glWinHeight
        
        '本文出力
        gobjExplorer.Document.Body.InnerHTML = _
            "<font face=""ＭＳ ゴシック"">" & _
            "<span style=""font-size:18px; line-height:22px;"">" & _
            gsProgMsg & "処理中...<br>" & _
            "　経過時間：" & ConvSec2SplitTime( Int( gdElapsedTime ) ) & "<br>" & _
            "　残り時間：" & ConvSec2SplitTime( RoundUp( gdRemainingTime ) ) & "<br>" & _
            String( gdProgPer10, "■") & String( 10 - gdProgPer10, "□") & "  " & gdProgPer100 & "% 完了" & _
            "</span>" & _
            "</font>" & _
            ""
    End Function
    
    ' ==================================================================
    ' = 概要    進捗値を変換する（例：100～500 を 0～1 に変換）
    ' = 引数    lInMin      Long   [in] 進捗最小値
    ' = 引数    lInMax      Long   [in] 進捗最大値
    ' = 引数    lInProg     Long   [in] 進捗値
    ' = 戻値                Double      変換結果（0 ～ 1 の小数値
    ' = 覚書    なし
    ' ==================================================================
    Public Function ConvProgRange( _
        ByVal lInMin, _
        ByVal lInMax, _
        ByVal lInProg _
    )
        Dim lConvMax
        Dim lConvProg
        
        If ( lInMin >= 0 And lInMax >= 0 And lInProg >= 0 ) And _
           ( lInMax >= lInProg And lInProg >= lInMin And lInMax >= lInMin ) Then
            'Do Nothing
        Else
            MsgBox "ConvProgRange関数の引数が不正です。" & vbNewLine & _
                   "lInMin  : " & lInMin & vbNewLine & _
                   "lInMax  : " & lInMax & vbNewLine & _
                   "lInProg : " & lInProg
            MsgBox "プログラムを中断します。"
            WScript.Quit()
        End If
        
        lConvMax = ( lInMax - lInMin ) + 1
        lConvProg = ( lInProg - lInMin ) + 1
        ConvProgRange = lConvProg / lConvMax
    End Function
    
    ' ==================================================================
    ' = 概要    プログレスバーを終了する
    ' = 引数    なし
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Function Quit()
        gobjExplorer.Document.Body.Style.Cursor = "default"
        gobjExplorer.Quit
    End Function
    
    Private Function ActiveIE()
        Dim objWshShell
        Dim intProcID
        
        Const strIEexe = "iexplore.exe" 'IEのプロセス名
        intProcID = GetProcID(strIEexe)
        
        Set objWshShell = CreateObject("Wscript.Shell")
        objWshShell.AppActivate intProcID
        
        Set objWshShell = Nothing
    End Function
    
    Private Function GetProcID( _
        ByVal ProcessName _
    )
        Dim Service
        Dim QfeSet
        Dim Qfe
        Dim intProcID
        
        Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
        Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='"& ProcessName &"'")
        
        intProcID = 0
        For Each Qfe in QfeSet
            intProcID = Qfe.ProcessId
            GetProcID = intProcID
            Exit For
        Next
    End Function
    
    Private Function ConvSec2SplitTime( _
        ByVal lRawSec _
    )
        Dim lOutSec
        Dim lOutMin
        Dim lOutHour
        Dim lMod
        
        lMod = lRawSec
        lOutHour = Fix(lMod / PROGBAR_ONEHOURSEC)
        lMod = Fix(lMod - (lOutHour * PROGBAR_ONEHOURSEC))
        lOutMin = Fix(lMod / PROGBAR_ONEMINSEC)
        lMod = Fix(lMod - (lOutMin * PROGBAR_ONEMINSEC))
        lOutSec = lMod
        
        If lOutHour = 0 Then
            If lOutMin = 0 Then
                ConvSec2SplitTime = lOutSec & " 秒"
            Else
                ConvSec2SplitTime = lOutMin & " 分 " & lOutSec & " 秒"
            End If
        Else
            ConvSec2SplitTime = lOutHour & " 時間 " & lOutMin & " 分 " & lOutSec & " 秒"
        End If
    End Function
    
    Private Function RoundUp( _
        ByVal dRawVal _
    )
        RoundUp = Round( dRawVal + 0.5 )
    End Function
    
    Private Function CalcWinHeight( _
        ByVal lLineNum _
    )
        CalcWinHeight = ( 28 * lLineNum ) + 65
    End Function
End Class
    If WScript.ScriptName = "ProgressBarIE.vbs" Then
        Call Test_ProgressBar
    End If
    Private Sub Test_ProgressBar
        Dim oProgBar
        Dim lTestCase
        Dim i
        Dim iBefore
        Dim iAfter
        Dim bTestContinue
        Dim bAllTestExec
        Dim bIsTestFinish
        
        lTestCase = InputBox( "テストケース番号を入力してください。" , "TestTitle" )
        If lTestCase = 0 Then
            bAllTestExec = True
        Else
            bAllTestExec = False
        End If
        
        bIsTestFinish = False
        bTestContinue = True
        Do While bTestContinue = True
            Set oProgBar = New ProgressBar
            Select Case lTestCase
                Case 0
                    'Do Nothing
                Case 1
                    oProgBar.Message = "Test Message"
                    iBefore = Timer()
                    For i = 0 to 100
                        oProgBar.Update( oProgBar.ConvProgRange( 0, 100, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case 2
                    oProgBar.Message = "Test Message"
                    iBefore = Timer()
                    For i = 400 to 500
                        oProgBar.Update( oProgBar.ConvProgRange( 400, 500, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case 3
                    oProgBar.Message = "Test Message"
                    WScript.Sleep 3000
                    iBefore = Timer()
                    For i = 0 to 1000
                        oProgBar.Update( oProgBar.ConvProgRange( 0, 1000, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case 4
                    oProgBar.Message = "Test Message" & vbNewLine & "aaa"
                    iBefore = Timer()
                    For i = 400 to 500
                        oProgBar.Update( oProgBar.ConvProgRange( 400, 500, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case 5
                    oProgBar.Message = "Test Message" & vbNewLine & "aaa" & vbNewLine & "aaa" & vbNewLine & "aaa" & vbNewLine & "aaa" & vbNewLine & "aaa" & vbNewLine & "aaa" & vbNewLine & "aaa"
                    iBefore = Timer()
                    For i = 400 to 500
                        oProgBar.Update( oProgBar.ConvProgRange( 400, 500, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case 6
                    iBefore = Timer()
                    For i = 0 to 1000
                        oProgBar.Update( oProgBar.ConvProgRange( 0, 1000, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case 7
                    iBefore = Timer()
                    For i = 0 to 1000
                        If i = 300 Then
                            oProgBar.Message = "Test Message" & vbNewLine & "ooo" & vbNewLine & "aaa" & vbNewLine & "aaa" & vbNewLine & "aaa" & vbNewLine & "aaa"
                        Else
                            'Do Nothing
                        End If
                        oProgBar.Update( oProgBar.ConvProgRange( 0, 1000, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case 8
                    oProgBar.Title = "Progress!!"
                    oProgBar.Message = "Test Message" & vbNewLine & "aaa"
                    iBefore = Timer()
                    For i = 0 to 100
                        oProgBar.Update( oProgBar.ConvProgRange( 0, 100, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case 9
                    iBefore = Timer()
                    oProgBar.Message = "Test Message" & vbNewLine & "aaa"
                    For i = 0 to 1000
                        oProgBar.Update( oProgBar.ConvProgRange( 0, 1000, i ) )
                        WScript.Sleep 10
                    Next
                    oProgBar.Message = "Test Message" & vbNewLine & "aaa" & vbNewLine & "aaa"
                    For i = 0 to 1000
                        oProgBar.Update( oProgBar.ConvProgRange( 0, 1000, i ) )
                        WScript.Sleep 10
                    Next
                    iAfter = Timer()
                    MsgBox iAfter - iBefore
                Case Else
                    bIsTestFinish = True
            End Select
            oProgBar.Quit()
            
            If bAllTestExec = True Then
                If bIsTestFinish = True Then
                    bTestContinue = False
                Else
                    lTestCase = lTestCase + 1
                    bTestContinue = True
                End If
            Else
                bTestContinue = False
            End If
        Loop
        Set oProgBar = Nothing
    End Sub

