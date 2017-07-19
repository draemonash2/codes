Attribute VB_Name = "Main_Read"
Option Explicit

Public Sub タグ情報読み込み()
    Dim vAnswer As Variant
    vAnswer = MsgBox("タグ読み込み処理を実行します。よろしいですか？", vbOKCancel)
    If vAnswer = vbOK Then
        'Do Nothing
    Else
        MsgBox "キャンセルされました"
        End
    End If
    
    Dim oProgBar As New ProgressBar
    Load oProgBar
    oProgBar.Show vbModeless
    
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    '################
    '### 共通処理 ###
    '################
    Call GetPreInfo
    Call ItunesInit
    
    Dim shTagList As Worksheet
    Set shTagList = ThisWorkbook.Sheets(TAG_LIST_SHEET_NAME)
    
    Dim lTagClmIdx As Long
    Dim lTagStrtClm As Long
    Dim lTagLastClm As Long
    lTagStrtClm = glRefStartClm + CLM_OFFSET_TAGINFO_TAGSTART
    lTagLastClm = shTagList.Cells(glRefStartRow + ROW_OFFSET_TITLE_02, shTagList.Columns.Count).End(xlToLeft).Column
    If lTagLastClm < lTagStrtClm Then
        MsgBox "読み書きするタグが指定されていません"
        End
    Else
        'Do Nothing
    End If
    
    Dim lRowIdx As Long
    Dim lStrtRow As Long
    Dim lLastRow As Long
    lStrtRow = glRefStartRow + ROW_OFFSET_TAG_START
    lLastRow = shTagList.Cells(shTagList.Rows.Count, glRefStartClm + CLM_OFFSET_TRACKINFO_FILEPATH).End(xlUp).Row
    If lLastRow < lStrtRow Then
        MsgBox "読み書きするタグが指定されていません"
        End
    Else
        'Do Nothing
    End If
    
    '############################################
    '### ミラーシート存在チェック＆追加＆整形 ###
    '############################################
    'ミラーシート追加
    Dim shSht As Worksheet
    Dim bIsShtExist As Boolean
    bIsShtExist = False
    For Each shSht In ThisWorkbook.Worksheets
        If shSht.Name = TAG_LIST_MIRROR_SHEET_NAME Then
            bIsShtExist = True
        Else
            'Do Nothing
        End If
    Next shSht
    Dim shTagListMir As Worksheet
    If bIsShtExist = True Then
        Set shTagListMir = ThisWorkbook.Sheets(TAG_LIST_MIRROR_SHEET_NAME)
    Else
        Set shTagListMir = ThisWorkbook.Sheets.Add(, ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        shTagListMir.Name = TAG_LIST_MIRROR_SHEET_NAME
    End If
    
    '値クリア
    shTagListMir.UsedRange.ClearContents
    
    'ミラーシートのタイトル部コピー
    For lTagClmIdx = lTagStrtClm To lTagLastClm
        shTagListMir.Cells(glRefStartRow + ROW_OFFSET_TITLE_01, lTagClmIdx).Value = shTagList.Cells(glRefStartRow + ROW_OFFSET_TITLE_01, lTagClmIdx).Value
        shTagListMir.Cells(glRefStartRow + ROW_OFFSET_TITLE_02, lTagClmIdx).Value = shTagList.Cells(glRefStartRow + ROW_OFFSET_TITLE_02, lTagClmIdx).Value
    Next lTagClmIdx
    
    'シート非表示
'   shTagListMir.Visible = False
    
    '##########################
    '### ログシート存在確認 ###
    '##########################
    bIsShtExist = False
    For Each shSht In ThisWorkbook.Worksheets
        If shSht.Name = ERROR_LOG_SHEET_NAME Then
            bIsShtExist = True
        Else
            'Do Nothing
        End If
    Next shSht
    Dim shLog As Worksheet
    If bIsShtExist = True Then
        Set shLog = ThisWorkbook.Sheets(ERROR_LOG_SHEET_NAME)
    Else
        MsgBox "シート「" & ERROR_LOG_SHEET_NAME & "」が見つかりません。"
        MsgBox "処理を中断します。"
        End
    End If
    Dim lLogRowIdx As Long
    Dim lLogLastRow As Long
    lLogLastRow = shLog.Cells(shLog.Rows.Count, 1).End(xlUp).Row + 1
    If lLogLastRow < LOG_START_ROW Then
        lLogLastRow = LOG_START_ROW
    Else
        'Do Nothing
    End If
    lLogRowIdx = lLogLastRow
    
    '##############################
    '### 「タグ一覧」シート更新 ###
    '##############################
    Dim sNow As String
    sNow = Now()
    
    'タグ情報領域クリア
    shTagList.Range( _
        shTagList.Cells(lStrtRow, lTagStrtClm), _
        shTagList.Cells(shTagList.Rows.Count, lTagLastClm) _
    ).ClearContents
    
    'タグ読み込み
    Dim sLogMsg As String
    Dim bIsErrorExist As Boolean
    bIsErrorExist = False
    Dim bIsTrackErrorExist As Boolean
    For lRowIdx = lStrtRow To lLastRow
        sLogMsg = "[Error]"
        bIsTrackErrorExist = False
        
        'トラック単位の読込対象確認
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            Dim sTrackExeEnable As String
            sTrackExeEnable = shTagList.Cells(lRowIdx, glRefStartClm + CLM_OFFSET_TRACKINFO_EXECUTE_ENABLE).Value
            Select Case sTrackExeEnable
                Case "○": 'Do Nothing
                Case "×": bIsTrackErrorExist = True
                Case "": 'Do Nothing
                Case Else: bIsTrackErrorExist = True
            End Select
            If bIsTrackErrorExist = True Then
                bIsErrorExist = True
                sLogMsg = sLogMsg & vbNewLine & "・読み込み対象トラックに指定されていません"
            Else
                'Do Nothing
            End If
        End If
        
        'ファイルパス取得
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            Dim sInFilePath As String
            sInFilePath = shTagList.Cells(lRowIdx, glRefStartClm + CLM_OFFSET_TRACKINFO_FILEPATH).Value
            If sInFilePath = "" Then
                bIsErrorExist = True
                bIsTrackErrorExist = True
                sLogMsg = sLogMsg & vbNewLine & "・ファイルパスが記載されていないため、ファイルが特定できません"
            Else
                Dim sFileExt As String
                sFileExt = LCase(ExtractTailWord(sInFilePath, "."))
                If sFileExt = "mp3" Then
                    'Do Nothing
                Else
                    bIsErrorExist = True
                    bIsTrackErrorExist = True
                    sLogMsg = _
                        sLogMsg & vbNewLine & _
                        "・mp3ファイルではありません。" & vbNewLine & _
                        "  sInFilePath : " & sInFilePath
                End If
                
            End If
        End If
        
        'トラック名取得
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            Dim bRet As Boolean
            Dim vFileInfoValue As Variant
            Dim vFileInfoTitle As Variant
            Dim sErrorDetail As String
            bRet = GetFileDetailInfo(sInFilePath, FILE_DETAIL_INFO_TRACK_NAME_INDEX, vFileInfoValue, vFileInfoTitle, sErrorDetail)
            If bRet = True Then
                If vFileInfoTitle = FILE_DETAIL_INFO_TRACK_NAME_TITLE Then
                    Dim sInTrackName As String
                    sInTrackName = CStr(vFileInfoValue)
                Else
                    Debug.Assert 0
                End If
            Else
                Select Case sErrorDetail
                    Case "File is not exist!": bIsTrackErrorExist = True
                    Case "Get info type error!": Debug.Assert 0
                    Case Else: Debug.Assert 0
                End Select
                If bIsTrackErrorExist = True Then
                    bIsErrorExist = True
                    sLogMsg = sLogMsg & vbNewLine & "・ファイルパスが存在しません"
                Else
                    'Do Nothing
                End If
            End If
        End If
        
        'トラック取得
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            Dim objTrack As Variant
            bRet = SearchTrack(sInTrackName, sInFilePath, objTrack)
            If bRet = True Then
                If objTrack.Location = sInFilePath Then
                    If objTrack Is Nothing Then
                        bIsErrorExist = True
                        bIsTrackErrorExist = True
                        sLogMsg = _
                            sLogMsg & vbNewLine & _
                            "・iTunes プレイリスト内にトラックがありません" & vbNewLine & _
                            "  sInFilePath : " & sInFilePath & vbNewLine & _
                            "  sInTrackName : " & sInTrackName
                    Else
                        'Do Nothing
                    End If
                Else
                    bIsErrorExist = True
                    bIsTrackErrorExist = True
                    sLogMsg = _
                        sLogMsg & vbNewLine & _
                        "・取得したトラック情報内のファイルパスが一致しません" & vbNewLine & _
                        "  sInFilePath : " & sInFilePath & vbNewLine & _
                        "  objTrack.Location : " & objTrack.Location
                End If
            Else
                bIsErrorExist = True
                bIsTrackErrorExist = True
                sLogMsg = _
                    sLogMsg & vbNewLine & _
                    "・iTunes プレイリスト内にトラックがありません" & vbNewLine & _
                    "  sInFilePath : " & sInFilePath & vbNewLine & _
                    "  sInTrackName : " & sInTrackName
            End If
        End If
        
        'タグ情報読み込み
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            For lTagClmIdx = lTagStrtClm To lTagLastClm
                Dim sTagTitle As String
                sTagTitle = shTagList.Cells(glRefStartRow + ROW_OFFSET_TITLE_02, lTagClmIdx).Value
                
                '読み込み処理
                Dim sTagValue As String
                bRet = GetTagValue(objTrack, sTagTitle, sTagValue)
                If bRet = True Then
                    shTagList.Cells(lRowIdx, lTagClmIdx).Value = sTagValue
                    shTagListMir.Cells(lRowIdx, lTagClmIdx).Value = sTagValue
                Else
                    bIsErrorExist = True
                    bIsTrackErrorExist = True
                    sLogMsg = _
                        sLogMsg & vbNewLine & _
                        "・指定されたタグタイトルのタグが見つかりませんでした" & vbNewLine & _
                        "  sTagTitle : " & sTagTitle
                End If
            Next lTagClmIdx
        End If
        
        If bIsTrackErrorExist = True Then
            shLog.Cells(lLogRowIdx, LOG_CLM_DATETIME).Value = sNow
            shLog.Cells(lLogRowIdx, LOG_CLM_RW).Value = "Read"
            shLog.Cells(lLogRowIdx, LOG_CLM_FILEPATH).Value = sInFilePath
            shLog.Cells(lLogRowIdx, LOG_CLM_TRACKNAME).Value = sInTrackName
            shLog.Cells(lLogRowIdx, LOG_CLM_ERRORMSG).Value = sLogMsg
            lLogRowIdx = lLogRowIdx + 1
        Else
            If OUTPUT_SUCCESS_LOG_TO_ERROR_LOG = True Then
                shLog.Cells(lLogRowIdx, LOG_CLM_DATETIME).Value = sNow
                shLog.Cells(lLogRowIdx, LOG_CLM_RW).Value = "Read"
                shLog.Cells(lLogRowIdx, LOG_CLM_FILEPATH).Value = sInFilePath
                shLog.Cells(lLogRowIdx, LOG_CLM_TRACKNAME).Value = sInTrackName
                shLog.Cells(lLogRowIdx, LOG_CLM_ERRORMSG).Value = "[Success]"
                lLogRowIdx = lLogRowIdx + 1
            Else
                'Do Nothing
            End If
        End If
        
        oProgBar.Update ((lRowIdx - lStrtRow) / (lLastRow - lStrtRow))
        
    Next lRowIdx
    
    Call ItunesTerminate
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    oProgBar.Hide
    Unload oProgBar
    
    If bIsErrorExist = True Then
        shLog.Activate
        MsgBox "エラーがあります！"
    Else
        MsgBox "タグ読み込みに成功しました！"
    End If
End Sub
