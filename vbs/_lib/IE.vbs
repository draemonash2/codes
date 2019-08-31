Option Explicit

' = 依存	なし
' = 所属	IE.vbs

' 定数は実験して導き出したもの
Const IE_LINE_HEIGHT_RATIO = 1.22
Const IE_WIN_LINE_HEIGHT_RATIO = 1.27
Const IE_HEADER_HEIGHT = 65

Class IE
    Dim gobjExplorer
    Dim glHorizontal
    Dim glVertical
    Dim gsFont
    Dim glFontSize
    Dim glLineHeight
    
    Private Sub Class_Initialize()
        '画面サイズ取得
        Dim objWMIService
        Dim colItems
        Dim objItem
        Set objWMIService = GetObject("Winmgmts:\\.\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
        For Each objItem in colItems
            glHorizontal = objItem.ScreenWidth
            glVertical = objItem.ScreenHeight
        Next
        Set objWMIService = Nothing
        Set colItems = Nothing
        
        gsFont = "ＭＳ ゴシック"
        glFontSize = 18
        glLineHeight = glFontSize * IE_LINE_HEIGHT_RATIO
        
        Set gobjExplorer = CreateObject("InternetExplorer.Application")
        gobjExplorer.Navigate "about:blank"
        gobjExplorer.ToolBar = 0
        gobjExplorer.StatusBar = 0
        gobjExplorer.Width = 450
        gobjExplorer.Height = 200
        gobjExplorer.Left = ( glHorizontal - gobjExplorer.Width ) / 2
        gobjExplorer.Top = ( glVertical - gobjExplorer.Height ) / 2
        gobjExplorer.Visible = 1
        
        gobjExplorer.Document.Body.InnerHTML = ""
    End Sub
    
    Private Sub Class_Terminate()
        ' Do Nothing
    End Sub
    
    ' ==================================================================
    ' = 概要    ウィンドウを作成、表示する
    ' = 引数    なし
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Sub Activate()
        gobjExplorer.Document.Body.Style.Cursor = "wait" 'マウスカーソルを砂時計にする
        Call ActiveIE
    End Sub
    
    ' ==================================================================
    ' = 概要    ウィンドウを終了する
    ' = 引数    なし
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Sub Quit()
        gobjExplorer.Document.Body.Style.Cursor = "default" 'マウスカーソルを元に戻す
        gobjExplorer.Quit
    End Sub
    
    ' ==================================================================
    ' = 概要    ウィンドウの高さを設定する
    ' = 引数    lHeight     Long   [in] ウィンドウの高さ
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    'ウィンドウのサイズはテキストの行数で自動算出するため、設定させない
'   Public Property Let Height( _
'       ByVal lHeight _
'   )
'       gobjExplorer.Height = lHeight
'       gobjExplorer.Top = ( glVertical - gobjExplorer.Height ) / 2
'   End Property
    
    ' ==================================================================
    ' = 概要    ウィンドウ幅を設定する
    ' = 引数    lWidth      Long   [in] ウィンドウの幅
    ' = 戻値    なし
    ' = 覚書    ウィンドウの表示位置は、幅に合わせて画面中央に自動調整する
    ' ==================================================================
    Public Property Let Width( _
        ByVal lWidth _
    )
        gobjExplorer.Width = lWidth
        gobjExplorer.Left = ( glHorizontal - gobjExplorer.Width ) / 2
    End Property
    
    ' ==================================================================
    ' = 概要    ウィンドウタイトルを設定する
    ' = 引数    sSetTitle   String   [in]   ウィンドウタイトル
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Property Let Title( _
        ByVal sSetTitle _
    )
        gobjExplorer.Document.Title = sSetTitle
    End Property
    
    ' ==================================================================
    ' = 概要    フォントを指定する
    ' = 引数    sFont   String   [in]   フォント名
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Property Let Font( _
        ByVal sFont _
    )
        gsFont = sFont
    End Property
    
    ' ==================================================================
    ' = 概要    フォントサイズを指定する
    ' = 引数    lFontSize   Long   [in] フォントサイズ
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Property Let FontSize( _
        ByVal lFontSize _
    )
        glFontSize = lFontSize
        glLineHeight = lFontSize * IE_LINE_HEIGHT_RATIO
    End Property
    
    ' ==================================================================
    ' = 概要    ウィンドウ内のテキストを設定する
    ' = 引数    sText   String   [in]   テキスト
    ' = 戻値    なし
    ' = 覚書    なし
    ' ==================================================================
    Public Property Let Text( _
        ByVal sText _
    )
        sText = Replace( sText, vbNewLine, "<br>" )
        sText = Replace( sText, vbCr, "<br>" )
        sText = Replace( sText, vbLf, "<br>" )
        
        'ウィンドウの高さ、位置算出
        Dim lLineNum
        lLineNum = ( ( Len( sText ) - Len( Replace( sText, "<br>", "" ) ) ) / 4 ) + 1
        gobjExplorer.Height = ( ( glLineHeight * IE_WIN_LINE_HEIGHT_RATIO ) * lLineNum ) + IE_HEADER_HEIGHT
        gobjExplorer.Top = ( glVertical - gobjExplorer.Height ) / 2
        
    '   MsgBox lLineNum & "：" & sText
        
        'テキスト設定
        gobjExplorer.Document.Body.InnerHTML = _
            "<font face=""" & gsFont & """>" & _
            "<span style=""font-size:" & glFontSize & "px; line-height:" & glLineHeight & "px;"">" & _
            sText & _
            "</span></font>"
    End Property
    
    Private Function ActiveIE()
        Dim Service
        Dim QfeSet
        Dim Qfe
        Dim lProcID
        Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
        Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='"& "iexplore.exe" &"'")
        lProcID = 0
        For Each Qfe in QfeSet
            lProcID = Qfe.ProcessId
            Exit For
        Next
        
        Dim objWshShell
        Set objWshShell = CreateObject("Wscript.Shell")
        objWshShell.AppActivate lProcID
        Set objWshShell = Nothing
    End Function
End Class
    If WScript.ScriptName = "IE.vbs" Then
        Call Test_IE
    End If
    Private Sub Test_IE
        Dim oIE
        Set oIE = New IE
        
        oIE.Activate
        WScript.Sleep(1000)
    '   Select Case 1
    '       Case 1:  oIE.Text = "あ"
    '       Case 2:  oIE.Text = "あ" & vbNewLine & "い"
    '       Case 3:  oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う"
    '       Case 4:  oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え"
    '       Case 8:  oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か" & vbNewLine & "あ" & vbNewLine & "い"
    '       Case 16: oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か" & vbNewLine & "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か"
    '       Case Else: MsgBox "error!"
    '   End Select
    
        oIE.Title = "タイトル"
        oIE.Font = "メイリオ"
        oIE.FontSize = 30
        oIE.Text = "あ"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か" & vbNewLine & "あ" & vbNewLine & "い"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か" & vbNewLine & "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か"
        WScript.Sleep(500)
        
        oIE.Font = "MS 明朝"
        oIE.FontSize = 8
        oIE.Text = "あ"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か" & vbNewLine & "あ" & vbNewLine & "い"
        WScript.Sleep(500)
        oIE.Text = "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か" & vbNewLine & "あ" & vbNewLine & "い" & vbNewLine & "う" & vbNewLine & "え" & vbNewLine & "お" & vbNewLine & "か"
        WScript.Sleep(500)
        
        oIE.Width = 500
        WScript.Sleep(1000)
        
        oIE.Quit
        Set oIE = Nothing
    End Sub
