Attribute VB_Name = "Main"
Option Explicit

Private Const RESULT_BOOK_NAME = "ChkResult.xlsx"

Private goTotalProcTimer As StopWatch

Private Function ChkExecInitialize()
    Set goTotalProcTimer = New StopWatch
    
    goTotalProcTimer.StartT '測定開始
    
'★デバッグ用    Application.ScreenUpdating = False
'★デバッグ用    Application.Calculation = xlCalculationManual
    
    Call ErrorMngInit
    Call InputInfoInit
    Call TcDocInfoInit
    Call ExcelFileInfoInit
    Call NewExcelMngInit
    Call SysInit
    Call SvnInit
    Call ChkTestDataInit
    Call ChkExistLogFileInit
    Call ChkTestDataOmissionInit
    Call ChkSummaryInit
End Function

Sub test222()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Function ChkExecTerminate()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Call TcDocInfoTerminate
    Unload PrgrsBar
    
    Debug.Print "==============================="
    Debug.Print "総経過時間          ：" & goTotalProcTimer.StopT & "[ms]" '測定停止
    Debug.Print ""
    Set goTotalProcTimer = Nothing
End Function

Public Function ChkExec()
    Dim wChkRsltBook As Workbook
    Dim sRsltBookPath As String
    Dim sFileListRootPath As String
    
    '##############
    '### 前処理 ###
    '##############
    Call ChkExecInitialize
    
    Debug.Print "初期化処理          ：" & goTotalProcTimer.IntervalTime & "[ms]"
    Load PrgrsBar
    PrgrsBar.Show vbModeless
    Call PrgrsBar.Update(10, "試験項目書取り込み中...")
    
    '##################
    '### メイン処理 ###
    '##################
    Call GetInputInfo
    Call GetTCDocInfo
    
    Debug.Print "項目書取込処理      ：" & goTotalProcTimer.IntervalTime & "[ms]"
    Call PrgrsBar.Update(40, "ファイルリスト作成中...")
    
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_ST Then
        sFileListRootPath = gtInputInfo.sTestLogDirPath & "\" & _
                            GetFileNameBase(gtInputInfo.sTestDocFilePath)
    Else
        sFileListRootPath = gtInputInfo.sTestLogDirPath & "\" & _
                            gtInputInfo.sSubjectName & "\" & _
                            GetFileNameBase(gtInputInfo.sTestDocFilePath)
    End If
    Call GetFileList(sFileListRootPath)
    
    Debug.Print "FileList作成処理    ：" & goTotalProcTimer.IntervalTime & "[ms]"
    Call PrgrsBar.Update(50, "チェック中...")
    
    Set wChkRsltBook = CreNewExcelFile
    
    Call ChkTestDataMain(wChkRsltBook)
    Call ChkExistLogFileMain(wChkRsltBook)
    Call ChkTestDataOmissionMain(wChkRsltBook)
    Call ChkSummaryMain(wChkRsltBook)
    
    Debug.Print "チェック処理        ：" & goTotalProcTimer.IntervalTime & "[ms]"
    Call PrgrsBar.Update(80, "コミット処理中...")
    
    sRsltBookPath = ThisWorkbook.Path & "\" & RESULT_BOOK_NAME 'チェック結果ファイル名
    sRsltBookPath = AddSeqNo2FilePath(sRsltBookPath) 'チェック結果ファイル名に連番「_000」を付与
    Call SaveNewExcelFile(wChkRsltBook, sRsltBookPath)
    
    Call ExecCommit
    
    Debug.Print "コミット処理        ：" & goTotalProcTimer.IntervalTime & "[ms]"
    Call PrgrsBar.Update(100)
    
    '##############
    '### 後処理 ###
    '##############
    Call ChkExecTerminate
End Function

Public Function WriTestRsltExec()
    Dim wWriRsltBook As Workbook
    Dim wTcDocBook As Workbook
    Dim sRsltBookPath As String
    Dim sFileListRootPath As String
    
    '##############
    '### 前処理 ###
    '##############
    Call ChkExecInitialize
    
    Debug.Print "初期化処理          ：" & goTotalProcTimer.IntervalTime & "[ms]"
    Load PrgrsBar
    PrgrsBar.Show vbModeless
    Call PrgrsBar.Update(10, "試験項目書取り込み中...")
    
    '##################
    '### メイン処理 ###
    '##################
    Call GetInputInfo
    Call GetTCDocInfo
    
    Debug.Print "項目書取込処理      ：" & goTotalProcTimer.IntervalTime & "[ms]"
    Call PrgrsBar.Update(50, "チェック中...")
    
    Set wWriRsltBook = CreNewExcelFile                           '書き込み結果ファイルオープン
    Set wTcDocBook = ExcelFileOpen(gtInputInfo.sTestDocFilePath) '項目書オープン
    Call CreBackupFile(gtInputInfo.sTestDocFilePath)             'バックアップファイル作成
    
    Call WriTestRslt(wTcDocBook, wWriRsltBook)
    
    Debug.Print "チェック処理        ：" & goTotalProcTimer.IntervalTime & "[ms]"
    Call PrgrsBar.Update(80, "コミット処理中...")
    
    Call ExcelFileClose(wTcDocBook, True)
    sRsltBookPath = ThisWorkbook.Path & "\" & RESULT_BOOK_NAME 'チェック結果ファイル名
    sRsltBookPath = AddSeqNo2FilePath(sRsltBookPath) 'チェック結果ファイル名に連番「_000」を付与
    Call SaveNewExcelFile(wWriRsltBook, sRsltBookPath)
    
    Call PrgrsBar.Update(100)
    
    '##############
    '### 後処理 ###
    '##############
    Call ChkExecTerminate
End Function

'★テスト用★
Private Function test()
    Dim lTcShtInfoCnt As Long
    Dim lTcInfoIdx As Long
    Dim lTestLogNameIdx As Long
    Dim vKey As Variant
    
    Debug.Print "### Input Info ###"
    Debug.Print "gtInputInfo.eTrgtPhase          = " & gtInputInfo.eTrgtPhase
    Debug.Print "gtInputInfo.sTestDocFilePath    = " & gtInputInfo.sTestDocFilePath
    Debug.Print "gtInputInfo.sTestLogDirPath     = " & gtInputInfo.sTestLogDirPath
    Debug.Print ""
    
    Debug.Print "### Test Case Doc Info ###"
    Debug.Print "gtTestDocInfo.eTrgtPhase          = " & gtTestDocInfo.eTrgtPhase
    Debug.Print "gtTestDocInfo.sTcDocName          = " & gtTestDocInfo.sTcDocName
    For Each vKey In gtTestDocInfo.oLogExpPathList
        Debug.Print "gtTestDocInfo.oLogExpPathList Key = " & vKey
    Next
    For lTcShtInfoCnt = 0 To UBound(gtTestDocInfo.atTcShtInfo)
        Debug.Print "gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).sModuleName  = " & gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).sModuleName
        Debug.Print "gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).sShtName     = " & gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).sShtName
        Debug.Print "gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).sSrcFileName = " & gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).sSrcFileName
        For lTcInfoIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo)
            Debug.Print "gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo(lTcInfoIdx).sTestCaseNo        = " & gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo(lTcInfoIdx).sTestCaseNo
            Debug.Print "gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo(lTcInfoIdx).sTestDataCellValue = " & gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo(lTcInfoIdx).sTestDataCellValue
            If Sgn(gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo(lTcInfoIdx).asTestLogName) = 0 Then
                'Do Nothing
            Else
                For lTestLogNameIdx = 0 To UBound(gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo(lTcInfoIdx).asTestLogName)
                    Debug.Print "gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo(lTcInfoIdx).asTestLogName (lTestLogNameIdx) = " & gtTestDocInfo.atTcShtInfo(lTcShtInfoCnt).atTcInfo(lTcInfoIdx).asTestLogName(lTestLogNameIdx)
                Next lTestLogNameIdx
            End If
        Next lTcInfoIdx
    Next lTcShtInfoCnt
    Debug.Print ""
    
'    Debug.Print "### File List Info ###"
'    For lFileListIdx = 0 To UBound(gatPathList)
'        Debug.Print "gatPathList(lFileListIdx).sPath     = [" & gatPathList(lFileListIdx).ePathType & "] " & gatPathList(lFileListIdx).sPath
'    Next lFileListIdx
'    Debug.Print ""
End Function
