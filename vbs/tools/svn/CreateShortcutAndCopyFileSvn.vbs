'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################
Const ADD_DATE_TYPE = 1 '�t�^��������̎�ʁi1:���ݓ����A2:�t�@�C��/�t�H���_�X�V�����j
Const SHORTCUT_FILE_SUFFIX = "#Src#"
Const ORIGINAL_FILE_PREFIX = "#Org#"
Const COPY_FILE_PREFIX     = "#Cpy#"

'####################################################################
'### ���O����
'####################################################################
Call Include( "C:\codes\vbs\_lib\String.vbs" ) 'ConvDate2String()

'####################################################################
'### �{����
'####################################################################
Const sPROG_NAME = "�V���[�g�J�b�g���R�s�[�t�@�C���쐬(SVN)"

Dim bIsContinue
bIsContinue = True

Dim sOrgDirPath
Dim cSelectedPaths
Dim objFSO

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Dim sArg
        Dim sDefaultPath
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        For Each sArg In WScript.Arguments
            cSelectedPaths.add sArg
            If sDefaultPath = "" Then
                sDefaultPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
        sOrgDirPath = InputBox( "�t�@�C���p�X���w�肵�Ă�������", sPROG_NAME, sDefaultPath )
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        sOrgDirPath = WScript.Env("Current")
        Set cSelectedPaths = WScript.Col( WScript.Env("Selected") )
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        sOrgDirPath = "X:\100_Documents\200_�y�w�Z�z����\��w�@\�[�~�o�ȕ�"
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        cSelectedPaths.Add "X:\100_Documents\200_�y�w�Z�z����\��w�@\�[�~�o�ȕ�\H20�N�x �[�~�o�ȕ�.xls"
        cSelectedPaths.Add "X:\100_Documents\200_�y�w�Z�z����\��w�@\�[�~�o�ȕ�\H21�N�x �[�~�o�ȕ�.xls"
    End If
Else
    'Do Nothing
End If

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
    If cSelectedPaths.Count = 0 Then
        MsgBox "�t�@�C��/�t�H���_���I������Ă��܂���B", vbOKOnly, sPROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �㏑���m�F ***
If bIsContinue = True Then
    Dim vbAnswer
    vbAnswer = MsgBox( "���Ƀt�@�C�������݂��Ă���ꍇ�A�㏑������܂��B���s���܂����H", vbOkCancel, sPROG_NAME )
    If vbAnswer = vbOk Then
        'Do Nothing
    Else
        MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

'*** �o�͐�I�� ***
If bIsContinue = True Then
    Dim sDstParDirPath
    sDstParDirPath = InputBox( "�o�͐����͂��Ă��������B", sPROG_NAME, sOrgDirPath )
    If sDstParDirPath = "" Then '�L�����Z���̏ꍇ
        MsgBox "���s���L�����Z������܂����B", vbOKOnly, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �V���[�g�J�b�g�쐬 ***
If bIsContinue = True Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    
    Dim sSelectedPath
    For Each sSelectedPath In cSelectedPaths
        '�t�@�C��/�t�H���_����
        Dim bFolderExists
        Dim bFileExists
        bFolderExists = objFSO.FolderExists( sSelectedPath )
        bFileExists = objFSO.FileExists( sSelectedPath )
        
        Dim sAddDate
        Dim sDstOrgFilePath
        Dim sDstCpyFilePath
        Dim sDstShrtctFilePath
        
        '### �t�@�C�� ###
        If bFolderExists = False And bFileExists = True Then
            '�ǉ�������擾�����`
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now(), 1 )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFile
                Set objFile = objFSO.GetFile( sSelectedPath )
                sAddDate = ConvDate2String( objFile.DateLastModified, 1 )
                Set objFile = Nothing
            Else
                MsgBox "�uADD_DATE_TYPE�v�̎w�肪����Ă��܂��I", vbOKOnly, sPROG_NAME
                MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
                Exit For
            End If
            
            Dim sOrgFileName
            Dim sOrgFileBaseName
            Dim sOrgFileExt
            sOrgFileName = objFSO.GetFileName( sSelectedPath )
            sOrgFileBaseName = objFSO.GetBaseName( sSelectedPath )
            sOrgFileExt = objFSO.GetExtensionName( sSelectedPath )
            sDstOrgFilePath    = sDstParDirPath & "\" & sOrgFileName & "_" & ORIGINAL_FILE_PREFIX & "_" & sAddDate & "." & sOrgFileExt
            sDstCpyFilePath    = sDstParDirPath & "\" & sOrgFileName & "_" & COPY_FILE_PREFIX     & "_" & sAddDate & "." & sOrgFileExt
            sDstShrtctFilePath = sDstParDirPath & "\" & sOrgFileName & "_" & SHORTCUT_FILE_SUFFIX & "_" & sAddDate & ".lnk"
            
            '�t�@�C���R�s�[
            objFSO.CopyFile sSelectedPath, sDstOrgFilePath, True
            objFSO.CopyFile sSelectedPath, sDstCpyFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### �t�H���_ ###
        ElseIf bFolderExists = True And bFileExists = False Then
            '�ǉ�������擾�����`
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now(), 1 )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFolder
                Set objFolder = objFSO.GetFolder( sSelectedPath )
                sAddDate = ConvDate2String( objFolder.DateLastModified, 1 )
                Set objFolder = Nothing
            Else
                MsgBox "�uADD_DATE_TYPE�v�̎w�肪����Ă��܂��I", vbOKOnly, sPROG_NAME
                MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
                Exit For
            End If
            
            Dim sOrgDirName
            sOrgDirName = objFSO.GetFileName( sSelectedPath )
            sDstOrgFilePath    = sDstParDirPath & "\" & sOrgDirName & "_" & ORIGINAL_FILE_PREFIX & "_" & sAddDate
            sDstCpyFilePath    = sDstParDirPath & "\" & sOrgDirName & "_" & COPY_FILE_PREFIX     & "_" & sAddDate
            sDstShrtctFilePath = sDstParDirPath & "\" & sOrgDirName & "_" & SHORTCUT_FILE_SUFFIX & "_" & sAddDate & ".lnk"
            
            '�t�H���_�R�s�[
            objFSO.CopyFolder sSelectedPath, sDstOrgFilePath, True
            objFSO.CopyFolder sSelectedPath, sDstCpyFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### �t�@�C��/�t�H���_�ȊO ###
        Else
            MsgBox "�I�����ꂽ�I�u�W�F�N�g�����݂��܂���" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, sPROG_NAME
            MsgBox "�����𒆒f���܂��B", vbOKOnly, sPROG_NAME
            bIsContinue = False
        End If
    Next
    
    Set objFSO = Nothing
    Set objWshShell = Nothing
    
    MsgBox "�V���[�g�J�b�g���R�s�[�t�@�C���̍쐬���������܂����I", vbOKOnly, sPROG_NAME
Else
    'Do Nothing
End If

'####################################################################
'### �C���N���[�h�֐�
'####################################################################
Private Function Include( ByVal sOpenFile )
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

