Attribute VB_Name = "Main_Write"
Option Explicit

Public Sub �^�O��񏑂�����()
    Dim vAnswer As Variant
    vAnswer = MsgBox("�^�O�������ݏ��������s���܂��B��낵���ł����H", vbOKCancel)
    If vAnswer = vbOK Then
        'Do Nothing
    Else
        MsgBox "�L�����Z������܂���"
        End
    End If
    
    Dim oProgBar As New ProgressBar
    Load oProgBar
    oProgBar.Show vbModeless
    
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    '################
    '### ���ʏ��� ###
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
        MsgBox "�ǂݏ�������^�O���w�肳��Ă��܂���"
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
        MsgBox "�ǂݏ�������^�O���w�肳��Ă��܂���"
        End
    Else
        'Do Nothing
    End If
    
    '################################
    '### �~���[�V�[�g���݃`�F�b�N ###
    '################################
    '�~���[�V�[�g�ǉ�
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
        MsgBox "�V�[�g�u" & TAG_LIST_MIRROR_SHEET_NAME & "�v������܂���B" & vbNewLine & "���O�Ƀ^�O����ǂݍ���ł�������"
        MsgBox "�����𒆒f���܂�"
        End
    End If
    
    '##########################
    '### ���O�V�[�g���݊m�F ###
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
        MsgBox "�V�[�g�u" & ERROR_LOG_SHEET_NAME & "�v��������܂���B"
        MsgBox "�����𒆒f���܂��B"
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
    
    '################
    '### �^�O�X�V ###
    '################
    '�~���[�V�[�g�̃^�C�g�����R�s�[
    For lTagClmIdx = lTagStrtClm To lTagLastClm
        shTagListMir.Cells(glRefStartRow + ROW_OFFSET_TITLE_01, lTagClmIdx).Value = shTagList.Cells(glRefStartRow + ROW_OFFSET_TITLE_01, lTagClmIdx).Value
        shTagListMir.Cells(glRefStartRow + ROW_OFFSET_TITLE_02, lTagClmIdx).Value = shTagList.Cells(glRefStartRow + ROW_OFFSET_TITLE_02, lTagClmIdx).Value
    Next lTagClmIdx
    
    Dim sNow As String
    sNow = Now()
    
    Dim sLogMsg As String
    Dim bIsErrorExist As Boolean
    bIsErrorExist = False
    Dim bIsTrackErrorExist As Boolean
    For lRowIdx = lStrtRow To lLastRow
        sLogMsg = "[Error]"
        bIsTrackErrorExist = False
        
        '�g���b�N�P�ʂ̓Ǎ��Ώۊm�F
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            Dim sTrackExeEnable As String
            sTrackExeEnable = shTagList.Cells(lRowIdx, glRefStartClm + CLM_OFFSET_TRACKINFO_EXECUTE_ENABLE).Value
            Select Case sTrackExeEnable
                Case "��": 'Do Nothing
                Case "�~": bIsTrackErrorExist = True
                Case "": 'Do Nothing
                Case Else: bIsTrackErrorExist = True
            End Select
            If bIsTrackErrorExist = True Then
                bIsErrorExist = True
                sLogMsg = sLogMsg & vbNewLine & "�E�������ݑΏۃg���b�N�Ɏw�肳��Ă��܂���"
            Else
                'Do Nothing
            End If
        End If
        
        '�t�@�C���p�X�擾
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            Dim sInFilePath As String
            sInFilePath = shTagList.Cells(lRowIdx, glRefStartClm + CLM_OFFSET_TRACKINFO_FILEPATH).Value
            If sInFilePath = "" Then
                bIsErrorExist = True
                bIsTrackErrorExist = True
                sLogMsg = sLogMsg & vbNewLine & "�E�t�@�C���p�X���L�ڂ���Ă��Ȃ����߁A�t�@�C��������ł��܂���"
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
                        "�Emp3�t�@�C���ł͂���܂���B" & vbNewLine & _
                        "  sInFilePath : " & sInFilePath
                End If
                
            End If
        End If
        
        '�g���b�N�擾
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            Dim lFileInfoTagIndex As Long
            Dim bRet
            bRet = GetFileDetailInfoIndex("�^�C�g��", lFileInfoTagIndex)
            If bRet = True Then
                'Do Nothing
            Else
                Debug.Assert 0
            End If
            
            Dim objTrack As Variant
            Dim sErrorDetail As String
            bRet = GetTrackInfo(sInFilePath, objTrack, sErrorDetail, lFileInfoTagIndex)
            If bRet = True Then
                'Do Nothing
            Else
                bIsErrorExist = True
                bIsTrackErrorExist = True
                If sErrorDetail = "File path is empty!" Then
                    Debug.Assert 0
                ElseIf sErrorDetail = "File is not exist at file system!" Then
                    sLogMsg = sLogMsg & vbNewLine & "�E�t�@�C���p�X�����݂��܂���"
                ElseIf sErrorDetail = "File is not exist at itunes playlist!" Then
                    sLogMsg = _
                        sLogMsg & vbNewLine & _
                        "�EiTunes �v���C���X�g���Ƀg���b�N������܂���" & vbNewLine & _
                        "  sInFilePath : " & sInFilePath
                Else
                    Debug.Assert 0
                End If
            End If
        End If
        
        '�^�O��񏑂�����
        If bIsTrackErrorExist = True Then
            'Do Nothing
        Else
            For lTagClmIdx = lTagStrtClm To lTagLastClm
                Dim sTagTitle As String
                sTagTitle = shTagList.Cells(glRefStartRow + ROW_OFFSET_TITLE_02, lTagClmIdx).Value
                
                '�������ݏ���
                Dim sTagValueMain As String
                Dim sTagValueMirror As String
                sTagValueMain = shTagList.Cells(lRowIdx, lTagClmIdx).Value
                sTagValueMirror = shTagListMir.Cells(lRowIdx, lTagClmIdx).Value
                If sTagValueMain = sTagValueMirror Then
                    'Do Nothing
                Else
                    bRet = SetTagValue(objTrack, sTagTitle, sTagValueMain)
                    If bRet = True Then
                        'Do Nothing
                    Else
                        '�������ݑΏۊO�̃^�O�͖���
                        
                        'bIsErrorExist = True
                        'bIsTrackErrorExist = True
                        'sLogMsg = _
                        '    sLogMsg & vbNewLine & _
                        '    "�E�w�肳�ꂽ�^�O�^�C�g���̃^�O��������܂���ł���" & vbNewLine & _
                        '    "  sTagTitle : " & sTagTitle
                    End If
                End If
            Next lTagClmIdx
        End If
        
        '�G���[��������
        If bIsTrackErrorExist = True Then
            shLog.Cells(lLogRowIdx, LOG_CLM_DATETIME).Value = sNow
            shLog.Cells(lLogRowIdx, LOG_CLM_RW).Value = "Write"
            shLog.Cells(lLogRowIdx, LOG_CLM_FILEPATH).Value = sInFilePath
            shLog.Cells(lLogRowIdx, LOG_CLM_ERRORMSG).Value = sLogMsg
            lLogRowIdx = lLogRowIdx + 1
        Else
            If OUTPUT_SUCCESS_LOG_TO_ERROR_LOG = True Then
                shLog.Cells(lLogRowIdx, LOG_CLM_DATETIME).Value = sNow
                shLog.Cells(lLogRowIdx, LOG_CLM_RW).Value = "Write"
                shLog.Cells(lLogRowIdx, LOG_CLM_FILEPATH).Value = sInFilePath
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
        MsgBox "�G���[������܂��I"
    Else
        MsgBox "�^�O�������݂ɐ������܂����I"
    End If
End Sub
