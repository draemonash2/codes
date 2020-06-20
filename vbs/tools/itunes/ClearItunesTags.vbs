Option Explicit

'==========================================================
'= �ݒ�l
'==========================================================
Const TRGT_DIR = "Z:\300_Musics"

Const DEBUG_FUNCVALID_BACKUPITUNELIBRARYS   = True
Const DEBUG_FUNCVALID_TAGUPDATE             = True

Const UPDATE_MODIFIED_DATE = False
Const ITUNES_BACKUP_FOLDER_MAX = 20

'==========================================================
'= �{����
'==========================================================
Dim objWshShell
Dim sCurDir
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
sCurDir = objWshShell.CurrentDirectory
Call Include( "C:\codes\vbs\_lib\String.vbs" )          'GetDirPath()
                                                        'GetFileName()
Call Include( "C:\codes\vbs\_lib\StopWatch.vbs" )       'class StopWatch
Call Include( "C:\codes\vbs\_lib\ProgressBarIE.vbs" )   'class ProgressBar
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )      'GetFileList2()
Call Include( "C:\codes\vbs\_lib\iTunes.vbs" )          '�����C���N���[�h����H
Call Include( "C:\codes\vbs\_lib\Array.vbs" )           '�����C���N���[�h����H

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
        "�@�EiTunes ���C�u���� �X�V����" & vbNewLine & _
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
' * ���C�u�����X�V                         *
' ******************************************
If DEBUG_FUNCVALID_TAGUPDATE = True Then ' ��Debug��
    
    objLogFile.WriteLine ""
    objLogFile.WriteLine "*** �^�O�X�V���� *** "
    oPrgBar.Message = _
        "�@�EiTunes ���C�u���� �o�b�N�A�b�v" & vbNewLine & _
        "�ˁEiTunes ���C�u���� �X�V����" & vbNewLine & _
        ""
    oPrgBar.Update( 0 ) '�i���X�V
    
    Dim oTrgtDirFiles
    Dim oTrgtFile
    Dim objPlayList
    Set objPlayList = WScript.CreateObject("iTunes.Application").Sources.Item(1).Playlists.ItemByName("�~���[�W�b�N")
    
    objLogFile.WriteLine "[Location]" & chr(9) & "[Grouping]" & chr(9) & "[Composer]"
    
    Dim lTrgtFileListNum
    Dim lTrgtFileListIdx
    lTrgtFileListNum = objPlayList.Tracks.Count
    lTrgtFileListIdx = 1
    
    Dim lHitFileNum
    lHitFileNum = 1
    
    Dim objTrack
    For Each objTrack In objPlayList.Tracks
        '�i���X�V
        oPrgBar.Update( _
            oPrgBar.ConvProgRange( _
                1, _
                lTrgtFileListNum, _
                lTrgtFileListIdx _
            ) _
        )
        
        If objTrack.KindAsString = "MPEG �I�[�f�B�I�t�@�C��" Then
            'If InStr( objTrack.Location, "Z:\300_Musics\290_Reggae@Riddim\Cardiac Keys" ) > 0 Then '��
            If objTrack.Grouping = "" And objTrack.Composer = "" Then
                'Do Nothing
            Else
                '�X�V���ޔ�
                If UPDATE_MODIFIED_DATE = True Then
                    '�X�V���͍X�V���ꂽ�܂�
                Else
                    Dim sTrgtFilePath
                    Dim sTrgtDirPath
                    Dim sTrgtFileName
                    Dim sTrgtModDate
                    sTrgtFilePath = objTrack.Location
                    sTrgtDirPath = GetDirPath(sTrgtFilePath)
                    sTrgtFileName = GetFileName(sTrgtFilePath)
                    Set oTrgtDirFiles = WScript.CreateObject("Shell.Application").Namespace( sTrgtDirPath & "\" )
                    Set oTrgtFile = oTrgtDirFiles.ParseName( sTrgtFileName )
                    sTrgtModDate = oTrgtFile.ModifyDate
                End If
                
                '�^�O�X�V
                objLogFile.WriteLine objTrack.Location & chr(9) & objTrack.Grouping & chr(9) & objTrack.Composer
                
                '�X�V����ύX���邽�߂Ɂu�W�������v�����������{�����߂�
                'objTrack.VolumeAdjustment = 0
                Dim sTestGenre
                sTestGenre = objTrack.Genre
                objTrack.Genre = "Test"
                'objTrack.Genre = sTestGenre
                
                objTrack.Grouping = ""
                objTrack.Composer = ""
                
                '�X�V�������߂�
                If UPDATE_MODIFIED_DATE = True Then
                    '�X�V���͍X�V���ꂽ�܂�
                Else
                    oTrgtFile.ModifyDate = CDate( sTrgtModDate ) '�X�V���������߂�
                End If
                lHitFileNum = lHitFileNum + 1
            End If
        Else
            'Do Nothing
        End If
        lTrgtFileListIdx = lTrgtFileListIdx + 1
    Next
    Set objPlayList = Nothing
    Set oTrgtFile = Nothing
    Set oTrgtDirFiles = Nothing
    
    objLogFile.WriteLine "�t�@�C�����F" & lHitFileNum
    objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime
    objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime
Else
    objLogFile.WriteLine "�o�ߎ��ԁi�{�����̂݁j : " & oStpWtch.IntervalTime
    objLogFile.WriteLine "�o�ߎ��ԁi�����ԁj     : " & oStpWtch.ElapsedTime
End If ' ��Debug��

' ******************************************
' * �I������                               *
' ******************************************
Call Finish
WScript.CreateObject("WScript.Shell").Run sLogFilePath, 1, True
MsgBox "�v���O����������ɏI�����܂����B"

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

'==========================================================
'= �C���N���[�h�֐�
'==========================================================
Private Function Include( _
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

