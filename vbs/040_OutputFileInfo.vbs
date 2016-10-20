Option Explicit

'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\lib\String.vbs" )

'==========================================================
'= 本処理
'==========================================================
Const INDEX_MAX = 500
'Const lContextLenBMax = 40

If WScript.Arguments.Count = 0 then
    MsgBox "情報を出力したいファイルを本スクリプトにドラッグ＆ドロップしてください。"
    MsgBox "プログラムを中断します。"
    WScript.Quit(-1)
Else
    Dim lArgIdx
    For lArgIdx = 0 to WScript.Arguments.Count - 1
        Dim sDirPath
        Dim sFileName
        Dim sFilePath
        sFilePath = WScript.Arguments( lArgIdx )
        sDirPath = GetDirPath( sFilePath )
        sFileName = GetFileName( sFilePath )
        
        Dim objFolder
        Dim objFile
        Set objFolder = CreateObject( "Shell.Application" ).Namespace( sDirPath )
        Set objFile = objFolder.ParseName( sFileName )
        
        Dim sLogPath
        sLogPath = sDirPath & "\" & Replace(sFileName, ".", "_") & ".txt"
'        sLogPath = sDirPath & "\" & Replace(Replace(sFileName," ", "_"), ".", "_") & ".txt"
        
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        On Error Resume Next
        Dim objLogFile
        Set objLogFile = objFSO.OpenTextFile( sLogPath, 2, True )
        If Err.Number <> 0 Then
            MsgBox Err.Number & "：" & Err.Description & vbNewLine & _
                   sLogPath
            WScript.Quit
        End If
        On Error Goto 0
        
        'MsgBox "指定されたファイルのファイル情報を以下に出力します。" & vbNewLine & _
        '      "  [ファイルパス] " & sLogPath & vbNewLine & _
        '       "  [文字コード] Unicode"
        
        Dim sItem
        Dim sContext
        
        '*** 項目数＆項目文字数算出 ***
        Dim lContextLenBMax
        Dim lIdx
        lContextLenBMax = 0
        For lIdx = 0 to INDEX_MAX
            sContext = objFolder.GetDetailsOf( objFolder.Items, lIdx )
            If sContext = "" Then
                'Do Nothing
            Else
                If Len( sContext ) > lContextLenBMax Then
                    lContextLenBMax = LenByte( sContext )
                Else
                    'Do Nothing
                End If
            End If
        Next
        
        '*** 項目出力 ***
        objLogFile.WriteLine "+-----+-" & String( lContextLenBMax , "-" ) & "-+-------------------------------------------------------"
        objLogFile.WriteLine "| idx | 項目名" & String( lContextLenBMax + 1 - LenByte("項目名"), " " ) & "| 値"
        objLogFile.WriteLine "+-----+-" & String( lContextLenBMax , "-" ) & "-+-------------------------------------------------------"
        
        Dim lContextNum
        lContextNum = 0
        For lIdx = 0 to INDEX_MAX
            sContext = objFolder.GetDetailsOf( objFolder.Items, lIdx )
            sItem = objFolder.GetDetailsOf( objFile, lIdx )
            
            If sContext = "" Or sItem = "" Then
                'Do Nothing
            Else
                On Error Resume Next
                Do
                    objLogFile.WriteLine "| " & String( 3 - Len(lIdx), " " ) & lIdx & " | " & _
                                          sContext & String( lContextLenBMax - LenByte(sContext), " " ) & " | " & _
                                          sItem
                    If Err.Number <> 0 Then
                        sItem = Right( sItem, Len(sItem) - 1 )
                        Err.Clear
                    Else
                        Exit Do
                    End If
                Loop While True
                On Error Goto 0
                lContextNum = lContextNum + 1
            End If
        Next
        objLogFile.WriteLine "+-----+-" & String( lContextLenBMax , "-" ) & "-+-------------------------------------------------------"
        objLogFile.WriteLine "【項目数】" & lContextNum
        objLogFile.Close
        
        Set objFolder = Nothing
        Set objFile = Nothing
        Set objFSO = Nothing
        Set objLogFile = Nothing
    Next
    MsgBox "完了！"
End if

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
    On Error Resume Next
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbNewLine & _
               sOpenFile
        WScript.Quit
    End If
    On Error Goto 0
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function
