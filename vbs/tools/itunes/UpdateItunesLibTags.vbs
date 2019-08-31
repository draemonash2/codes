Option Explicit

'����������������������������������������������������������������������������������������
'��
'�� iTunes �^�O�X�V�c�[��
'��
'��  �y�T�v�z
'��     mp3 �t�@�C���̍X�V���������Ƀ^�O�X�V�ς݃t�@�C�����������ʂ��AiTunes �� API ��
'��     �R�}���h���C��������s���āAiTunes ���C�u������ mp3 �^�O�����X�V����B
'��     
'��     iTunes �ȊO�̃\�t�g�� mp3 �^�O���X�V�����ꍇ�AiTunes ���N�����Ă��Â��^�O��
'��     �܂ܕ\������A���C�u�����͍X�V����Ȃ��B���̎��A�Y���̃t�@�C�����Đ����邩�A
'��     iTunes ��ŉ��炩�̃^�O���X�V����΃��C�u�����X�V����邪�A��ʂ̃t�@�C����
'��     �X�V�����ꍇ�A���Ɏ�Ԃ�������B�{�c�[���́A���̎�Ԃ��Ȃ����Ƃ�ړI�Ƃ���
'��     �쐬�����B
'��     
'��     �܂��A�w��t�H���_�z���̃��C�u�������o�^�̃t�@�C���������I�ɓo�^�ł���B
'��     �i�o�^���s�ۂ͑I���j
'��     
'��  �y���ӎ����A���L�����z
'��     �E�w�肵���t�H���_�z���Ɏ��s���ʃ��O���i�[����B
'��     �E�u��Ȏҁv�u�O���[�v�v�u�R�����g�v�i���ɂ����邩���j�ɂ����Ă� iTunes �ȊO��
'��       �\�t�g�ōX�V���Ă��AiTunes ��̕\���͍X�V����Ȃ��B�iiTunes �̎d�l�j 
'��       �\�����X�V���邽�߂ɂ́AiTunes ���璼�ڏ�L�^�O���X�V����K�v������B
'��       �{�c�[�������l�̗��R�ŁA��L�^�O�̕\���X�V�͂ł��Ȃ��B
'��     
'��  �y�g�p���@�z
'��     (1) ���y���i�[����t�H���_�̃��[�g�t�H���_�p�X���uTRGT_DIR�v�ɋL�ڂ���B
'��     (2) �uTRGT_DIR�v�ɋL�ڂ���B
'��     (3) �{�X�N���v�g�����s�B
'��     
'��  �y�X�V�����z
'��     v2.8 (2017/03/21)
'��       �E���������Ώۃ^�O��ύX
'��           Composer �� VolumeAdjustment
'��       �E�X�V�������߂������ύX
'��     
'��     v2.7 (2017/03/15)
'��       �E�Â� iTunes Library Backup �t�H���_�̍폜
'��     
'��     v2.6 (2017/03/05)
'��       �EStopWatch.vbs �C���ɑ΂���Ή�
'��       �E�y���ȃo�OFix
'��     
'��     v2.5 (2016/10/28)
'��       �E���O�t�@�C���o�͐�ύX
'��       �E�����I�����Ƀ��O�t�@�C�����J��������ǉ�
'��     
'��     v2.4 (2016/10/27)
'��       �E�����I��ǉ�
'��     
'��     v2.3 (2016/10/18)
'��       �EiTunes ���C�u�����o�b�N�A�b�v�@�\�ǉ�
'��     
'��     v2.2 (2016/10/17)
'��       �E����
'��     
'����������������������������������������������������������������������������������������

'==========================================================
'= �ݒ�l
'==========================================================
Const TRGT_DIR = "Z:\300_Musics"
Const UPDATE_MODIFIED_DATE = False
Const ITUNES_BACKUP_FOLDER_MAX = 20

Const DEBUG_FUNCVALID_BACKUPITUNELIBRARYS   = True
Const DEBUG_FUNCVALID_ADDFILES              = True
Const DEBUG_FUNCVALID_DATEINPUT             = True
Const DEBUG_FUNCVALID_TRGTLISTUP            = True
Const DEBUG_FUNCVALID_DIRCMDEXEC            = True
Const DEBUG_FUNCVALID_TAGUPDATE             = True
Const DEBUG_FUNCVALID_DIRRESULTDELETE       = True

'==========================================================
'= �{����
'==========================================================
Dim objWshShell
Dim sCurDir
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
sCurDir = objWshShell.CurrentDirectory
Call Include( "C:\codes\vbs\_lib\String.vbs" )          'GetDirPath()
                                                        'RemoveTailWord()
                                                        'ExtractTailWord()
Call Include( "C:\codes\vbs\_lib\StopWatch.vbs" )       'class StopWatch
Call Include( "C:\codes\vbs\_lib\ProgressBarIE.vbs" )   'class ProgressBar
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )      'GetFileList2()
Call Include( "C:\codes\vbs\_lib\iTunes.vbs" )          '�����C���N���[�h����H
Call Include( "C:\codes\vbs\_lib\Array.vbs" )           '�����C���N���[�h����H

' ******************************************
' * �����I��                               *
' ******************************************
Dim bIsExecLibAdd
Dim bIsExecLibMod
Dim sAnswer
sAnswer = MsgBox( "iTunes �փt�@�C����ǉ����܂����H" & vbNewLine & _
                  "  [�ǉ��Ώۃt�H���_] " & TRGT_DIR _
                  , vbYesNoCancel _
                )
Select Case sAnswer
    Case vbYes: bIsExecLibAdd = True
    Case vbNo:  bIsExecLibAdd = False
    Case Else:
        MsgBox "�������L�����Z�����܂����B" & vbNewLine & _
               "�v���O�������I�����܂��B"
        WScript.Quit
End Select
sAnswer = MsgBox( "�o�^�ς݂̋Ȃɂ��� iTunes ���C�u�����̃^�O���X�V���܂����H" & vbNewLine & _
                  "  [�X�V�Ώۃt�H���_] " & TRGT_DIR _
                  , vbYesNoCancel _
                )
Select Case sAnswer
    Case vbYes: bIsExecLibMod = True
    Case vbNo:  bIsExecLibMod = False
    Case Else:
        MsgBox "�������L�����Z�����܂����B" & vbNewLine & _
               "�v���O�������I�����܂��B"
        WScript.Quit
End Select
sAnswer = MsgBox( "���s���鏈���͈ȉ��Ŗ�肠��܂��񂩁H" & vbNewLine & _
                  "  �EiTunes ���C�u�����ǉ�         �F " & bIsExecLibAdd & vbNewLine & _
                  "  �EiTunes ���C�u�����X�V         �F " & bIsExecLibMod _
                  , vbYesNo _
                )
Select Case sAnswer
    Case vbYes:
        'Do Nothing
    Case vbNo:
        MsgBox "�������L�����Z�����܂����B" & vbNewLine & _
               "�v���O�������I�����܂��B"
        WScript.Quit
End Select

' ******************************************
' * ���O����                               *
' ******************************************
If 1 Then '�����u���b�N���̂��߂̕��򏈗�
    '*** �X�g�b�v�E�H�b�`�N�� ***
    Dim oStpWtch
    Set oStpWtch = New StopWatch
    Call oStpWtch.StartT
    
    '*** �v���O���X�o�[�N�� ***
    Dim oPrgBar
    Set oPrgBar = New ProgressBar
End If

' ******************************************
' * iTunes ���C�u�����o�b�N�A�b�v          *
' ******************************************
If DEBUG_FUNCVALID_BACKUPITUNELIBRARYS = True Then ' ��Debug��
    oPrgBar.Message = _
        "�ˁEiTunes ���C�u���� �o�b�N�A�b�v" & vbNewLine & _
        "�@�EiTunes ���C�u���� �ǉ�����" & vbNewLine & _
        "�@�EiTunes ���C�u���� �X�V����" & vbNewLine & _
        "�@�@- ���t���͏���" & vbNewLine & _
        "�@�@- �X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
        "�@�@- �^�O�X�V����" & _
        ""
    oPrgBar.Update( 0.2 ) '�i���X�V
    
    '*** �o�b�N�A�b�v�t�H���_�쐬 ***
    Dim sCurDateTime
    sCurDateTime = Now()
    
    Dim objItunes
    Set objItunes = WScript.CreateObject("iTunes.Application")
    
    Dim sBackUpDirName
    Dim sBackUpDirPath
    Dim sItuneDirPath
    Dim sItuneBackUpDirPath
    sItuneDirPath = GetDirPath( objItunes.LibraryXMLPath )
    sItuneBackUpDirPath = sItuneDirPath & "\iTunes Library Backup"
    sBackUpDirName = Year( sCurDateTime ) & _
                     String( 2 - Len( Month( sCurDateTime ) ), "0" ) & Month( sCurDateTime ) & _
                     String( 2 - Len( Day( sCurDateTime ) ), "0" ) & Day( sCurDateTime ) & _
                     "_" & _
                     String( 2 - Len( Hour( sCurDateTime ) ), "0" ) & Hour( sCurDateTime ) & _
                     String( 2 - Len( Minute( sCurDateTime ) ), "0" ) & Minute( sCurDateTime ) & _
                     String( 2 - Len( Second( sCurDateTime ) ), "0" ) & Second( sCurDateTime )
    sBackUpDirPath = sItuneBackUpDirPath & "\" & sBackUpDirName
    
    Set objItunes = Nothing
    
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder( sBackUpDirPath )
    objFSO.CopyFile sItuneDirPath & "\iTunes Library Extras.itdb", sBackUpDirPath & "\"
    objFSO.CopyFile sItuneDirPath & "\iTunes Library Genius.itdb", sBackUpDirPath & "\"
    objFSO.CopyFile sItuneDirPath & "\iTunes Library.itl        ", sBackUpDirPath & "\"
    objFSO.CopyFile sItuneDirPath & "\iTunes Music Library.xml  ", sBackUpDirPath & "\"
    
    '*** �Â��o�b�N�A�b�v�t�H���_���폜 ***
    Dim asDirList
    Call GetFileList2(sItuneBackUpDirPath, asDirList, 2) 
    
    '�t�H���_�폜
    If UBound( asDirList ) >= ITUNES_BACKUP_FOLDER_MAX then
        Dim lDelFolderMax
        lDelFolderMax = UBound(asDirList) - ITUNES_BACKUP_FOLDER_MAX
        Dim lDelDirIdx
        For lDelDirIdx = LBound(asDirList) to lDelFolderMax
            '�o�b�N�A�b�v�t�H���_���́uYYYYMMDD_HHMMSS�v�œ��ꂳ��Ă��邽�߁A
            'asDirList() �͎��R�Ɠ������ɕ��ԁB�i�v�f�ԍ����傫���Ȃ�قǐV�����j
            '���̂��߁A�v�f�ԍ��̏�����������t�H���_���폜����B
            objFSO.DeleteFolder asDirList(lDelDirIdx), True
        Next
    Else
        'Do Nothing
    End If
    
    '*** ���O�t�@�C���쐬 ***
    Dim sLogFilePath
    sLogFilePath = sBackUpDirPath & "\" & Replace( WScript.ScriptName, ".vbs", ".log" )
    
    Dim objLogFile
    Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )
    
    objLogFile.WriteLine "script started."
    objLogFile.WriteLine ""
    objLogFile.WriteLine "[�X�V�Ώۃt�H���_] " & TRGT_DIR
    objLogFile.WriteLine ""
    objLogFile.WriteLine "*** iTunes ���C�u�����o�b�N�A�b�v *** "
    objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime
    objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime
Else
    sLogFilePath = TRGT_DIR & "\" & Replace( WScript.ScriptName, ".vbs", ".log" )
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objLogFile = objFSO.OpenTextFile( sLogFilePath, 2, True )
    
    objLogFile.WriteLine "script started."
    objLogFile.WriteLine ""
    objLogFile.WriteLine "[�X�V�Ώۃt�H���_] " & TRGT_DIR
End If ' ��Debug��

' ******************************************
' * �t�@�C���ǉ�                           *
' ******************************************
if bIsExecLibAdd = True Then
    
    If DEBUG_FUNCVALID_ADDFILES = True Then ' ��Debug��
        
        objLogFile.WriteLine ""
        objLogFile.WriteLine "*** �t�@�C���ǉ� *** "
        oPrgBar.Message = _
            "�@�EiTunes ���C�u���� �o�b�N�A�b�v" & vbNewLine & _
            "�ˁEiTunes ���C�u���� �ǉ�����" & vbNewLine & _
            "�@�EiTunes ���C�u���� �X�V����" & vbNewLine & _
            "�@�@- ���t���͏���" & vbNewLine & _
            "�@�@- �X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
            "�@�@- �^�O�X�V����" & _
            ""
        oPrgBar.Update( 0.5 ) '�i���X�V
        
        WScript.CreateObject("iTunes.Application").LibraryPlaylist.AddFile( TRGT_DIR )
        
        oPrgBar.Update( 1 ) '�i���X�V
        
        objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime
        objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime
    Else
        'Do Nothing
    End If ' ��Debug��
Else
    'Do Nothing
End If

' ******************************************
' * ���C�u�����X�V                         *
' ******************************************
if bIsExecLibMod = True Then
    ' ==============================
    ' = ���t����                   =
    ' ==============================
    If DEBUG_FUNCVALID_DATEINPUT = True Then ' ��Debug��
        
        objLogFile.WriteLine ""
        objLogFile.WriteLine "*** ���t���͏��� *** "
        oPrgBar.Message = _
            "�@�EiTunes ���C�u���� �o�b�N�A�b�v" & vbNewLine & _
            "�@�EiTunes ���C�u���� �ǉ�����" & vbNewLine & _
            "�@�EiTunes ���C�u���� �X�V����" & vbNewLine & _
            "�ˁ@- ���t���͏���" & vbNewLine & _
            "�@�@- �X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
            "�@�@- �^�O�X�V����" & _
            ""
        oPrgBar.Update( 0 ) '�i���X�V
        
        On Error Resume Next
        
        oPrgBar.Update( 0.1 ) '�i���X�V
        
        Dim sNow
        sNow = Now()
        sNow = Left( sNow, Len( sNow ) - 2 ) & "00" '�b��00�ɂ���
        
        Dim sCmpBaseTime
        sCmpBaseTime = InputBox( _
                            "�X�V�ΏۂƂ���t�@�C������肵�܂��B" & vbNewLine & _
                            "�X�V�ΏۂƂ��鎞������͂��Ă��������B" & vbNewLine & _
                            "" & vbNewLine & _
                            "  [���͋K��] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
                            "" & vbNewLine & _
                            "�� ���t�݂̂��w�肵�����ꍇ�A�uYYYY/MM/DD 0:0:0�v�Ƃ��Ă��������B" _
                            , "����" _
                            , sNow _
                        )
        
        objLogFile.WriteLine "���͒l : " & sCmpBaseTime
        
        Dim sTimeValue
        Dim sDateValue
        sTimeValue = TimeValue(sCmpBaseTime)
        sDateValue = DateValue(sCmpBaseTime)
        
        oPrgBar.Update( 0.5 ) '�i���X�V
        
        '���t�`�F�b�N
        If Err.Number <> 0 Then
            MsgBox "���t�̌`�����s���ł��I" & vbNewLine & _
                   "  [���͋K��] YYYY/MM/DD HH:MM:SS" & vbNewLine & _
                   "  [���͒l] " & sCmpBaseTime
            MsgBox Err.Description
            MsgBox "�v���O�����𒆒f���܂��I"
            Err.Clear
            Call Finish
            WScript.Quit
        Else
            'Do Nothing
        End If
        If DateDiff("s", sCmpBaseTime, Now() ) < 0  Then
            MsgBox "�����̓������w�肳��܂����I" & vbNewLine & _
                   "  [���͒l] " & sCmpBaseTime
            MsgBox "�v���O�����𒆒f���܂��I"
            Call Finish
            WScript.Quit
        Else
            'Do Nothing
        End If
        On Error Goto 0 '�uOn Error Resume Next�v������
        
        oPrgBar.Update( 1 ) '�i���X�V
        
        oStpWtch.IntervalTime ' IntervalTime �X�V
        
    Else ' ��Debug��
        sCmpBaseTime = "2016/10/27 10:00:00"
    End If ' ��Debug��
    
    ' ==============================
    ' = �X�V�Ώۃt�@�C�����X�g�擾 =
    ' ==============================
    If DEBUG_FUNCVALID_TRGTLISTUP = True Then ' ��Debug��
        
        objLogFile.WriteLine ""
        objLogFile.WriteLine "*** �X�V�Ώۃt�@�C������ *** "
        oPrgBar.Message = _
            "�@�EiTunes ���C�u���� �o�b�N�A�b�v" & vbNewLine & _
            "�@�EiTunes ���C�u���� �ǉ�����" & vbNewLine & _
            "�@�EiTunes ���C�u���� �X�V����" & vbNewLine & _
            "�@�@- ���t���͏���" & vbNewLine & _
            "�ˁ@- �X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
            "�@�@- �^�O�X�V����" & _
            ""
        oPrgBar.Update( 0 ) '�i���X�V
        
        On Error Resume Next
        
        '*** Dir �R�}���h���s ***
        Dim sTmpFilePath
        Dim sExecCmd
        sTmpFilePath = objWshShell.CurrentDirectory & "\" & replace( WScript.ScriptName, ".vbs", "_TrgtFileList.tmp" )
        If DEBUG_FUNCVALID_DIRCMDEXEC = True Then ' ��Debug��
            sExecCmd = "Dir """ & TRGT_DIR & """ /s /a:a-d > """ & sTmpFilePath & """"
            With CreateObject("Wscript.Shell")  
                .Run "cmd /c" & sExecCmd, 7, True
            End With
        End If ' ��Debug��
        
        '*** Dir �R�}���h���ʎ擾 ***
        Dim objFile
        Dim sTextAll
        If Err.Number = 0 Then
            Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
            If Err.Number = 0 Then
                sTextAll = objFile.ReadAll
                sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '�����ɉ��s���t�^����Ă��܂����߁A�폜
                objFile.Close
            Else
                WScript.Echo "�t�@�C�����J���܂���: " & Err.Description
            End If
            Set objFile = Nothing   '�I�u�W�F�N�g�̔j��
        Else
            WScript.Echo "�G���[ " & Err.Description
        End If
        On Error Goto 0
        
        oPrgBar.Update( 0.2 ) '�i���X�V
        
        '*** �X�V�������o ***
        Dim oMatchResult
        Dim sSearchPattern
        Dim oRegExp
        Dim sTargetStr
        Set oRegExp = CreateObject("VBScript.RegExp")
        sSearchPattern = "((\d{4}/\d{1,2}/\d{1,2})\s+(\d{1,2}:\d{1,2})\s+([0-9,]+)\s+(.+)\r)|(\s+(.*)\s�̃f�B���N�g��)"
        sTargetStr = sTextAll
        oRegExp.Pattern = sSearchPattern               '�����p�^�[����ݒ�
        oRegExp.IgnoreCase = True                      '�啶���Ə���������ʂ��Ȃ�
        oRegExp.Global = True                          '������S�̂�����
        Set oMatchResult = oRegExp.Execute(sTargetStr) '�p�^�[���}�b�`���s
        
        Dim sFileName
        Dim sFilePath
        Dim sFileSize
        Dim sModDate
        Dim sDirName
        Dim iMatchIdx
        Dim sExtName
        Dim asTrgtFileList()
        ReDim asTrgtFileList(-1)
        Dim lMatchResultCount
        sDirName = ""
        objLogFile.WriteLine "[sFilePath]" & chr(9) & _
                             "[sDirName]"  & chr(9) & _
                             "[sFileName]" & chr(9) & _
                             "[sModDate]"  & chr(9) & _
                             "[sFileSize]" & chr(9) & _
                             "[sExtName]"
        lMatchResultCount = oMatchResult.Count
        For iMatchIdx = 0 To lMatchResultCount - 1
            '�i���X�V
            If iMatchIdx Mod 100 = 0 Then
                oPrgBar.Update( _
                    oPrgBar.ConvProgRange( _
                        0, _
                        oMatchResult.Count - 1, _
                        iMatchIdx _
                    ) _
                )
            End If
            If oMatchResult(iMatchIdx).SubMatches.Count = 7 Then
                '�t�@�C�����Ƀ}�b�`
                If oMatchResult(iMatchIdx).SubMatches(0) <> "" Then
                    sModDate = oMatchResult(iMatchIdx).SubMatches(1) & " " & _
                               oMatchResult(iMatchIdx).SubMatches(2)
                    sFileSize = oMatchResult(iMatchIdx).SubMatches(3)
                    sFileName = oMatchResult(iMatchIdx).SubMatches(4)
                    sFilePath = sDirName & "\" & sFileName
                    sExtName = ExtractTailWord( sFileName, "." )
                    
        '           objLogFile.WriteLine oMatchResult(iMatchIdx).SubMatches(0) & chr(9) & _
        '                                oMatchResult(iMatchIdx).SubMatches(1) & chr(9) & _
        '                                oMatchResult(iMatchIdx).SubMatches(2) & chr(9) & _
        '                                oMatchResult(iMatchIdx).SubMatches(3) & chr(9) & _
        '                                oMatchResult(iMatchIdx).SubMatches(4) & chr(9) & _
        '                                oMatchResult(iMatchIdx).SubMatches(5) & chr(9) & _
        '                                oMatchResult(iMatchIdx).SubMatches(6)
                    
                    '�X�V������r �� �X�V�ΏۑI��
                    If LCase(sExtName) = "mp3" Then
                        If DateDiff("s", sCmpBaseTime, sModDate ) >= 0  Then
                            ReDim Preserve asTrgtFileList( UBound(asTrgtFileList) + 1 )
                            asTrgtFileList( UBound(asTrgtFileList) ) = sFilePath
                            objLogFile.WriteLine sFilePath & chr(9) & _
                                                 sDirName  & chr(9) & _
                                                 sFileName & chr(9) & _
                                                 sModDate  & chr(9) & _
                                                 sFileSize & chr(9) & _
                                                 sExtName
                        Else
                            'Do Nothing
                        End If
                    Else
                        'Do Nothing
                    End If
                    
                '�t�H���_���Ƀ}�b�`
                ElseIf oMatchResult(iMatchIdx).SubMatches(5) <> "" Then
                    sDirName = oMatchResult(iMatchIdx).SubMatches(6)
                Else
                    MsgBox "�G���[�I"
                End If
            Else
                MsgBox "�G���[�I"
            End If
        Next
        
        oPrgBar.Update( 1 ) '�i���X�V
        
        If DEBUG_FUNCVALID_DIRRESULTDELETE = True Then ' ��Debug��
            objFSO.DeleteFile sTmpFilePath, True
        End If ' ��Debug��
        
        Set objFSO = Nothing    '�I�u�W�F�N�g�̔j��
        
        objLogFile.WriteLine "�t�@�C�����F" & UBound(asTrgtFileList) + 1
        objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime
        objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime
        
    Else ' ��Debug��
        ReDim asTrgtFileList(0)
        asTrgtFileList(0) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Bow Down.mp3"
    '   asTrgtFileList(1) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concentrate.mp3"
    '   asTrgtFileList(2) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Concrete Schoolyard.mp3"
    '   asTrgtFileList(3) = "Z:\300_Musics\600_HipHop\Artist\$ Other\Control Myself.mp3"
    End If ' ��Debug��
    
    ' ==============================
    ' = �^�O�X�V                   =
    ' ==============================
    If DEBUG_FUNCVALID_TAGUPDATE = True Then ' ��Debug��
        
        objLogFile.WriteLine ""
        objLogFile.WriteLine "*** �^�O�X�V���� *** "
        oPrgBar.Message = _
            "�@�EiTunes ���C�u���� �o�b�N�A�b�v" & vbNewLine & _
            "�@�EiTunes ���C�u���� �ǉ�����" & vbNewLine & _
            "�@�EiTunes ���C�u���� �X�V����" & vbNewLine & _
            "�@�@- ���t���͏���" & vbNewLine & _
            "�@�@- �X�V�Ώۃt�@�C�����菈��" & vbNewLine & _
            "�ˁ@- �^�O�X�V����" & _
            ""
        oPrgBar.Update( 0 ) '�i���X�V
        
        objLogFile.WriteLine "[TrgtFileIdx}" & Chr(9) & "[HitIdx] / [HitNum}" & Chr(9) & "[FilePath}" & Chr(9) & "[TrackName]" & Chr(9) & "[Kind]" & Chr(9) & "[Location]" & Chr(9) & "[LocMatch]"
        
        Dim lTrgtFileListIdx
        Dim lTrgtFileListNum
        lTrgtFileListNum = UBound( asTrgtFileList )
        For lTrgtFileListIdx = 0 To lTrgtFileListNum
            '�i���X�V
            oPrgBar.Update( _
                oPrgBar.ConvProgRange( _
                    0, _
                    lTrgtFileListNum, _
                    lTrgtFileListIdx _
                ) _
            )
            
            Dim sTrgtFilePath
            sTrgtFilePath = asTrgtFileList( lTrgtFileListIdx )
            
            '�g���b�N���擾
            Dim sTrgtDirPath
            Dim sTrgtFileName
            sTrgtDirPath = RemoveTailWord( sTrgtFilePath, "\" )
            sTrgtFileName = ExtractTailWord( sTrgtFilePath, "\" )
            
            Dim oTrgtDirFiles
            Dim oTrgtFile
            Dim sTrgtTrackName
            Dim sTrgtModDate
            Set oTrgtDirFiles = CreateObject("Shell.Application").Namespace( sTrgtDirPath )
            Set oTrgtFile = oTrgtDirFiles.ParseName( sTrgtFileName )
            sTrgtTrackName = oTrgtDirFiles.GetDetailsOf( oTrgtFile, 21 )
            sTrgtModDate = oTrgtFile.ModifyDate
            
            Dim objPlayList
            Dim objSearchResult
            Set objPlayList = WScript.CreateObject("iTunes.Application").Sources.Item(1).Playlists.ItemByName("�~���[�W�b�N")
            Set objSearchResult = objPlayList.Search( sTrgtTrackName, 5 )
            
            Dim sOutLine
            Dim bIsFilePathMatched
            Dim lHitIdx
            bIsFilePathMatched = False
            For lHitIdx = 1 to objSearchResult.Count
                With objSearchResult.Item(lHitIdx)
                    sOutLine = ( lTrgtFileListIdx + 1 ) & Chr(9) & lHitIdx & " / " & objSearchResult.Count & Chr(9) & sTrgtFilePath & Chr(9) & sTrgtTrackName & Chr(9) & .Kind
                    If .Kind = 1 Then
                        sOutLine = sOutLine & Chr(9) & .Location & Chr(9) & ( .Location = sTrgtFilePath )
                        If .Location = sTrgtFilePath Then
                            .VolumeAdjustment = 0
                            bIsFilePathMatched = True
                        Else
                            'Do Nothing
                        End If
                    Else
                        'Do Nothing
                    End If
                End With
                objLogFile.WriteLine sOutLine
                If bIsFilePathMatched = True Then
                    Exit For
                Else
                    'Do Nothing
                End If
            Next
            Set objSearchResult = Nothing
            Set objPlayList = Nothing
            
            If UPDATE_MODIFIED_DATE = True Then
                '�X�V���͍X�V���ꂽ�܂�
            Else
                oTrgtFile.ModifyDate = CDate( sTrgtModDate ) '�X�V���������߂�
            End If
            
            Set oTrgtFile = Nothing
            Set oTrgtDirFiles = Nothing
        Next
        
        objLogFile.WriteLine "�t�@�C�����F" & UBound(asTrgtFileList) + 1
        objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime
        objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime
        
    End If ' ��Debug��
Else
    'Do Nothing
End If

' ******************************************
' * �I������                               *
' ******************************************
Call Finish
WScript.CreateObject("WScript.Shell").Run "%comspec% /c """ & sLogFilePath & """"
MsgBox "�v���O����������ɏI�����܂����B"

'==========================================================
'= �֐���`
'==========================================================
' �O���v���O���� �C���N���[�h�֐�
Function Include( _
    ByVal sOpenFile _
)
    Dim objFSO
    Dim objVbsFile
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function

Function Finish()
    Call oStpWtch.StopT
    Call oPrgBar.Quit
    objLogFile.WriteLine ""
    objLogFile.WriteLine "�J�n����               : " & oStpWtch.StartPoint
    objLogFile.WriteLine "�I������               : " & oStpWtch.StopPoint
    objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime
    objLogFile.WriteLine ""
    objLogFile.WriteLine "script finished."
    objLogFile.Close
    Set oStpWtch = Nothing
    Set oPrgBar = Nothing
End Function
