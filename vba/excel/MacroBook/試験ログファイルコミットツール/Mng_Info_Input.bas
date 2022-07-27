Attribute VB_Name = "Mng_Info_Input"
Option Explicit

Private Const INPUT_SRCH_KEY_TRGT_PHASE = "試験フェーズ"
Private Const INPUT_SRCH_KEY_SBJCT_NAME = "案件名"
Private Const INPUT_SRCH_KEY_DOC_PATH = "試験項目書ファイルパス"
Private Const INPUT_SRCH_KEY_LOG_PATH = "試験ログフォルダパス"
Private Const INPUT_SRCH_KEY_TESTER = "評価者"
Private Const INPUT_SRCH_KEY_TEST_DATE = "年月日"
Private Const INPUT_SRCH_KEY_TEST_RSLT = "結果判定"
Private Const INPUT_SRCH_KEY_TEST_DATA = "試験データ"
Private Const INPUT_SRCH_KEY_REV_SRC = "試験 Rev（ソースコード）"
Private Const INPUT_SRCH_KEY_REV_HEXABS = "試験 Rev（HEX/ABS）"
Private Const INPUT_SRCH_KEY_REV_A2L = "試験 Rev（A2L）"

Private Const INPUT_SHEET_NAME = "データ入力"

Private Const TRGT_PHASE_NAME_UT = "単体試験"
Private Const TRGT_PHASE_NAME_CT = "結合試験"
Private Const TRGT_PHASE_NAME_FT = "機能試験"
Private Const TRGT_PHASE_NAME_ST = "システム試験"

Public Enum E_TRGT_PHASE
    TRGT_PHASE_UT
    TRGT_PHASE_CT
    TRGT_PHASE_FT
    TRGT_PHASE_ST
End Enum

Public Type T_INPUT_INFO
    eTrgtPhase As E_TRGT_PHASE
    sSubjectName As String
    sTestDocFilePath As String
    sTestLogDirPath As String
    sTester As String
    sTestDate As String
    sTestRslt As String
    sRevSrc As String
    sRevHexAbs As String
    sRevA2L As String
End Type

Public gtInputInfo As T_INPUT_INFO

Public Function InputInfoInit()
    Dim tInputInfoInit As T_INPUT_INFO
    gtInputInfo = tInputInfoInit
End Function

Public Function GetInputInfo()
    Dim shTrgtSht As Worksheet
    Dim sFileBaseName As String
    Dim sTrgtPhaseName As String
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    Set shTrgtSht = ThisWorkbook.Sheets(INPUT_SHEET_NAME)
    
    '### セルデータ取得 ###
    '*** 試験フェーズ ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_TRGT_PHASE, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("「" & INPUT_SRCH_KEY_DOC_PATH & "」が記載されておりません！")
        Call OutpErrorMsg(ERROR_PROC_STOP)
    Else
        'Do Nothing
    End If
    sTrgtPhaseName = tNearCellData.sCellValue
    Select Case sTrgtPhaseName
        Case TRGT_PHASE_NAME_UT:    gtInputInfo.eTrgtPhase = TRGT_PHASE_UT
        Case TRGT_PHASE_NAME_CT:    gtInputInfo.eTrgtPhase = TRGT_PHASE_CT
        Case TRGT_PHASE_NAME_FT:    gtInputInfo.eTrgtPhase = TRGT_PHASE_FT
        Case TRGT_PHASE_NAME_ST:    gtInputInfo.eTrgtPhase = TRGT_PHASE_ST
        Case Else:                  Stop
    End Select
    
    '*** 案件名 ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_SBJCT_NAME, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("「" & INPUT_SRCH_KEY_SBJCT_NAME & "」が記載されておりません！")
    Else
        'Do Nothing
    End If
    gtInputInfo.sSubjectName = tNearCellData.sCellValue
    
    '*** 試験項目書ファイルパス ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_DOC_PATH, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("「" & INPUT_SRCH_KEY_DOC_PATH & "」が記載されておりません！")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTestDocFilePath = tNearCellData.sCellValue
    
    '*** 試験ログフォルダパス ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_LOG_PATH, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("「" & INPUT_SRCH_KEY_LOG_PATH & "」が記載されておりません！")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTestLogDirPath = tNearCellData.sCellValue
    
    '*** 評価者 ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_TESTER, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("「" & INPUT_SRCH_KEY_TESTER & "」が記載されておりません！")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTester = tNearCellData.sCellValue
    
    '*** 年月日 ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_TEST_DATE, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("「" & INPUT_SRCH_KEY_TEST_DATE & "」が記載されておりません！")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTestDate = tNearCellData.sCellValue
    
    '*** 結果判定 ***
    tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_TEST_RSLT, 0, 1)
    If tNearCellData.bIsCellDataExist = False Then
        Stop
    Else
        'Do Nothing
    End If
    If tNearCellData.sCellValue = "" Then
        Call StoreErrorMsg("「" & INPUT_SRCH_KEY_TEST_RSLT & "」が記載されておりません！")
    Else
        'Do Nothing
    End If
    gtInputInfo.sTestRslt = tNearCellData.sCellValue
    
    '*** 試験データ ***
    '試験データは格納しない
    
    If gtInputInfo.eTrgtPhase = TRGT_PHASE_UT Then
        '*** 試験 Rev（ソースコード） ***
        tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_REV_SRC, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        If tNearCellData.sCellValue = "" Then
            Call StoreErrorMsg("「" & INPUT_SRCH_KEY_REV_SRC & "」が記載されておりません！")
        Else
            'Do Nothing
        End If
        gtInputInfo.sRevSrc = tNearCellData.sCellValue
    Else
        '*** 試験 Rev（HEX/ABS） ***
        tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_REV_HEXABS, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        If tNearCellData.sCellValue = "" Then
            Call StoreErrorMsg("「" & INPUT_SRCH_KEY_REV_HEXABS & "」が記載されておりません！")
        Else
            'Do Nothing
        End If
        gtInputInfo.sRevHexAbs = tNearCellData.sCellValue
        
        '*** 試験 Rev（A2L） ***
        tNearCellData = GetNearCellData(shTrgtSht, INPUT_SRCH_KEY_REV_A2L, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        If tNearCellData.sCellValue = "" Then
            Call StoreErrorMsg("「" & INPUT_SRCH_KEY_REV_A2L & "」が記載されておりません！")
        Else
            'Do Nothing
        End If
        gtInputInfo.sRevA2L = tNearCellData.sCellValue
    End If
    
    Call OutpErrorMsg(ERROR_PROC_STOP)
    
    '### データチェック処理 ###
    '試験項目書名⇔試験フェーズ一致チェック
    sFileBaseName = GetFileNameBase(gtInputInfo.sTestDocFilePath)
    If InStr(sFileBaseName, sTrgtPhaseName) > 0 Then
        'Do Nothing
    Else
        Call StoreErrorMsg("項目書名に「" & sTrgtPhaseName & "」が含まれておりません！")
        Call OutpErrorMsg(ERROR_PROC_THROUGH)
    End If
End Function

Public Function OutputDocFilePathCell( _
    ByVal sPath As String _
)
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    
    With ThisWorkbook
        '試験項目書ファイルパス
        tNearCellData = GetNearCellData(.Sheets(INPUT_SHEET_NAME), INPUT_SRCH_KEY_DOC_PATH, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        
        .Sheets(INPUT_SHEET_NAME).Cells(tNearCellData.lRow, tNearCellData.lClm).Value = sPath
    End With
End Function

Public Function OutputLogDirPathCell( _
    ByVal sPath As String _
)
    Dim tNearCellData As T_EXCEL_NEAR_CELL_DATA
    
    With ThisWorkbook
        '試験項目書ファイルパス
        tNearCellData = GetNearCellData(.Sheets(INPUT_SHEET_NAME), INPUT_SRCH_KEY_LOG_PATH, 0, 1)
        If tNearCellData.bIsCellDataExist = False Then
            Stop
        Else
            'Do Nothing
        End If
        
        .Sheets(INPUT_SHEET_NAME).Cells(tNearCellData.lRow, tNearCellData.lClm).Value = sPath
    End With
End Function



