Option Explicit

'<<�T�v>>
'  �w�肵���t�@�C�����o�b�N�A�b�v����B
'  
'<<�g�p���@>>
'  BackUpFile.vbs <filepath> <backupnum> <logfilepath>
'  
'<<�d�l>>
'  �E�t�@�C�����w�肷��ƌ��ݎ�����t�^�����o�b�N�A�b�v�t�@�C�����쐬����B
'  �E�����t�@�C�����̂��̂����݂��Ă�����A�A���t�@�x�b�g��t�^�����o�b�N�A�b�v�t�@�C�����쐬����B
'     ex. 211201a, 211202b, �c
'  �E�������Ɏw�肳�ꂽ�o�b�N�A�b�v�����t�@�C�������܂�����A�Â����̂���폜����B
'  �E���s���ʂ͑�O�����Ɏw�肳�ꂽ���O�t�@�C���ɏo�͂���B
'  �E�O��o�b�N�A�b�v������X�V����Ă��Ȃ��ꍇ�A�o�b�N�A�b�v���Ȃ��B
'  
'<<���ӎ���>>
'  �E�o�b�N�A�b�v�Ώۂ̓t�@�C���̂݁B
'  �E�ȉ���S�Ė������ꍇ�A�V�����t�@�C�����X�V����Ă������ߗv���ӁB
'      - �o�b�N�A�b�v�t�@�C���̐ڔ�����"z"�ƂȂ��Ă���t�@�C�������� (ex. file_#b#211122z.txt)
'  �E�ȉ��̗��R�ōŐV/�ŌÃo�b�N�A�b�v�t�@�C������ɍX�V������p���Ȃ��B�����܂�
'    �o�b�N�A�b�v�������������t�@�C�����Ŕ��f����B
'      ����ČÂ��o�b�N�A�b�v�t�@�C�����X�V���Ă��܂����ꍇ�A�t�@�C�������
'      ���t���Â��̂ɍX�V�������V�����t�@�C�����ł��Ă��܂��B
'      �X�V���������Ƃɔ��肷��ƁA��L�̃t�@�C�����폜���ꂸ�A�c���Ă��܂����߁B

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" ) 'GetFileListCmdClct()
                                                            'CreateDirectry()
                                                            'GetFileInfo()

'===============================================================================
'= �ݒ�l
'===============================================================================
Const sBAK_DIR_NAME = "_bak"
Const sBAK_FILE_SUFFIX = "#b#"

'===============================================================================
'= �{����
'===============================================================================
Const sSCRIPT_NAME = "�t�@�C���o�b�N�A�b�v"
Dim sBakSrcFilePath
Dim lBakFileNumMax
Dim sBakLogFilePath
If WScript.Arguments.Count >= 3 Then
    sBakSrcFilePath = WScript.Arguments(0)
    lBakFileNumMax = CLng(WScript.Arguments(1))
    sBakLogFilePath = WScript.Arguments(2)
Else
    WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
    WScript.Quit
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objLogFile
Set objLogFile = objFSO.OpenTextFile(sBakLogFilePath, 8, True) 'AddWrite

'****************
'*** ���O���� ***
'****************
'�Ώۃt�@�C�����擾
Dim sBakSrcParDirPath
Dim sBakSrcFileBaseName
Dim sBakSrcFileExt
Dim sDateSuffix
sBakSrcParDirPath = objFSO.GetParentFolderName( sBakSrcFilePath )
sBakSrcFileBaseName = objFSO.GetBaseName( sBakSrcFilePath )
sBakSrcFileExt = objFSO.GetExtensionName( sBakSrcFilePath )
sDateSuffix = ConvDate2String(Now(),2)

'�o�b�N�A�b�v�t�@�C�����쐬
Dim sBakDstDirPath
Dim sBakDstPathBase
sBakDstDirPath = sBakSrcParDirPath & "\" & sBAK_DIR_NAME
sBakDstPathBase = sBakDstDirPath & "\" & sBakSrcFileBaseName & "_" & sBAK_FILE_SUFFIX

'****************************
'*** �t�@�C���o�b�N�A�b�v ***
'****************************
'�o�b�N�A�b�v�t�H���_�쐬
Call CreateDirectry( sBakDstDirPath )

'�t�@�C���ꗗ�擾
Dim cFileList
Set cFileList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct( sBakDstDirPath, cFileList, 1, "*")

'�����̍ŐV�t�@�C���T��
Dim sBakDstFilePathLatest  '�����̍ŐV�o�b�N�A�b�v�t�@�C��
sBakDstFilePathLatest = ""
Dim sFilePath
For Each sFilePath In cFileList
    If ( ( InStr(sFilePath, sBakDstPathBase) > 0 ) And _
       (objFSO.GetExtensionName(sFilePath) = sBakSrcFileExt) ) Then
        sBakDstFilePathLatest = sFilePath
    End If
Next
Set cFileList = Nothing

'�o�b�N�A�b�v�t�@�C�����m��
Dim sBakDstFilePath
'�����̃o�b�N�A�b�v�t�@�C�������݂��A�������t�̃o�b�N�A�b�v�t�@�C�������݂���ꍇ
If sBakDstFilePathLatest <> "" And _
   InStr(sBakDstFilePathLatest, sBakDstPathBase & sDateSuffix) > 0 Then
    Dim sTailChar
    sTailChar = Right( objFSO.GetBaseName( sBakDstFilePathLatest ), 1)
    Dim lBakDstAlphaIdx
    If Asc(sTailChar) >= Asc("a") And Asc(sTailChar) < Asc("z") Then
        lBakDstAlphaIdx = Asc(sTailChar) + 1
    ElseIf Asc(sTailChar) = Asc("z") Then
        lBakDstAlphaIdx = Asc(sTailChar)
    ElseIf Asc(sTailChar) >= Asc("0") And Asc(sTailChar) <= Asc("9") Then
        lBakDstAlphaIdx = Asc("a")
    Else
        objLogFile.WriteLine "�s���ȃo�b�N�A�b�v�t�@�C����������܂����B"
        objLogFile.WriteLine "  " & sBakDstFilePathLatest
        objLogFile.WriteLine "�v���O�����𒆒f���܂��B"
        WScript.Quit
    End If
    sBakDstFilePath = sBakDstPathBase & sDateSuffix & Chr(lBakDstAlphaIdx) & "." & sBakSrcFileExt
Else
    sBakDstFilePath = sBakDstPathBase & sDateSuffix & "." & sBakSrcFileExt
End If
'objLogFile.WriteLine sBakDstFilePath & " : " & sBakDstFilePathLatest
'WScript.Quit

'�X�V�����擾
Dim vDateLastModifiedLatestBk
Dim vDateLastModifiedTrgt
Dim bRet
bRet = GetFileInfo( sBakDstFilePathLatest, 11, vDateLastModifiedLatestBk)
bRet = GetFileInfo( sBakSrcFilePath, 11, vDateLastModifiedTrgt)

'�����̃o�b�N�A�b�v�t�@�C�������� or �X�V����Ă���ꍇ
If ( sBakDstFilePathLatest = "" ) Or _
   ( ( sBakDstFilePathLatest <> "" ) And ( vDateLastModifiedTrgt > vDateLastModifiedLatestBk ) ) Then
    '�t�@�C���o�b�N�A�b�v
    objFSO.CopyFile sBakSrcFilePath, sBakDstFilePath, True
    objLogFile.WriteLine "[Success] " & sBakSrcFilePath & " -> " & sBakDstFilePath
Else
    '�O��o�b�N�A�b�v������X�V����Ă��Ȃ��ꍇ�A�o�b�N�A�b�v���������𒆒f����
    objLogFile.WriteLine "[Skip]    " & sBakSrcFilePath
    WScript.Quit
End If

'************************
'*** �Â��t�@�C���폜 ***
'************************
'�t�@�C�����X�g�擾
Set cFileList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct( sBakDstDirPath, cFileList, 1, "*")

'�o�b�N�A�b�v�t�@�C�����擾�{�����̍ŌÃt�@�C���T��
Dim lBakFileNum
Dim sDelFilePath
lBakFileNum = 0
sDelFilePath = ""
For Each sFilePath in cFileList
    If ( (InStr(sFilePath, sBakDstPathBase) > 0) And _
         (objFSO.GetExtensionName(sFilePath) = sBakSrcFileExt) ) Then
        If lBakFileNum = 0 Then
           sDelFilePath = sFilePath
        End If
        lBakFileNum = lBakFileNum + 1
    End If
Next

'�o�b�N�A�b�v�t�@�C���폜
If lBakFileNum > lBakFileNumMax Then
    objFSO.DeleteFile sDelFilePath, True
End If

'objLogFile.WriteLine "�o�b�N�A�b�v�����I", vbOKOnly, sSCRIPT_NAME

objLogFile.Close

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function
