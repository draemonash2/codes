'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################
Const ADD_DATE_TYPE = 1 '�t�^��������̎�ʁi1:���ݓ����A2:�t�@�C��/�t�H���_�X�V�����j
Const EVACUATE_ORG_FILE = True
Const SHORTCUT_FILE_SUFFIX = "#s#"
Const ORIGINAL_FILE_PREFIX = "#o#"
Const EDIT_FILE_PREFIX     = "#e#"

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�V���[�g�J�b�g���R�s�[�t�@�C���쐬"

Dim bIsContinue
bIsContinue = True

Dim sOrgDirPath
Dim cSelectedPaths
Dim objFSO

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Dim sArg
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set cSelectedPaths = CreateObject("System.Collections.ArrayList")
        For Each sArg In WScript.Arguments
            cSelectedPaths.add sArg
            If sOrgDirPath = "" Then
                sOrgDirPath = objFSO.GetParentFolderName( sArg )
            End If
        Next
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
        MsgBox "�t�@�C��/�t�H���_���I������Ă��܂���B", vbOKOnly, PROG_NAME
        MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �㏑���m�F ***
'���s���x�����߂邽�߁A�㏑���m�F�ȗ�
'If bIsContinue = True Then
'    Dim vbAnswer
'    vbAnswer = MsgBox( "���Ƀt�@�C�������݂��Ă���ꍇ�A�㏑������܂��B���s���܂����H", vbOkCancel, PROG_NAME )
'    If vbAnswer = vbOk Then
'        'Do Nothing
'    Else
'        MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
'        bIsContinue = False
'    End If
'Else
'    'Do Nothing
'End If

'*** �o�͐�I�� ***
If bIsContinue = True Then
    Dim sDstParDirPath
    sDstParDirPath = ShowFolderSelectDialog( sOrgDirPath )

    If objFSO.FolderExists( sDstParDirPath ) = False Then '�L�����Z���̏ꍇ
        MsgBox "���s���L�����Z������܂����B", vbOKOnly, PROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �ޔ�p�t�H���_�쐬 ***
If bIsContinue = True Then
    Dim sDstParEvaDirPath
    If EVACUATE_ORG_FILE = True Then
        sDstParEvaDirPath = sDstParDirPath & "\_" & ORIGINAL_FILE_PREFIX
        '�t�H���_�쐬
        If objFSO.FolderExists( sDstParEvaDirPath ) = False Then
            objFSO.CreateFolder( sDstParEvaDirPath )
        End If
    Else
        sDstParEvaDirPath = sDstParDirPath
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
                sAddDate = ConvDate2String( Now() )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFile
                Set objFile = objFSO.GetFile( sSelectedPath )
                sAddDate = ConvDate2String( objFile.DateLastModified )
                Set objFile = Nothing
            Else
                MsgBox "�uADD_DATE_TYPE�v�̎w�肪����Ă��܂��I", vbOKOnly, PROG_NAME
                MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
                Exit For
            End If
            
            Dim sSrcFileName
            Dim sSrcFileBaseName
            Dim sSrcFileExt
            sSrcFileName = objFSO.GetFileName( sSelectedPath )
            sSrcFileBaseName = objFSO.GetBaseName( sSelectedPath )
            sSrcFileExt = objFSO.GetExtensionName( sSelectedPath )
            sDstCpyFilePath    = sDstParDirPath    & "\" & sSrcFileName & "_" & EDIT_FILE_PREFIX     & sAddDate & "." & sSrcFileExt
            sDstOrgFilePath    = sDstParEvaDirPath & "\" & sSrcFileName & "_" & ORIGINAL_FILE_PREFIX & sAddDate & "." & sSrcFileExt
            sDstShrtctFilePath = sDstParEvaDirPath & "\" & sSrcFileName & "_" & SHORTCUT_FILE_SUFFIX & sAddDate & ".lnk"
            
            '�t�@�C���R�s�[
            objFSO.CopyFile sSelectedPath, sDstCpyFilePath, True
            objFSO.CopyFile sSelectedPath, sDstOrgFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### �t�H���_ ###
        ElseIf bFolderExists = True And bFileExists = False Then
            '�ǉ�������擾�����`
            If ADD_DATE_TYPE = 1 Then
                sAddDate = ConvDate2String( Now() )
            ElseIf ADD_DATE_TYPE = 2 Then
                Dim objFolder
                Set objFolder = objFSO.GetFolder( sSelectedPath )
                sAddDate = ConvDate2String( objFolder.DateLastModified )
                Set objFolder = Nothing
            Else
                MsgBox "�uADD_DATE_TYPE�v�̎w�肪����Ă��܂��I", vbOKOnly, PROG_NAME
                MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
                Exit For
            End If
            
            Dim sSrcDirName
            sSrcDirName = objFSO.GetFileName( sSelectedPath )
            sDstCpyFilePath    = sDstParDirPath    & "\" & sSrcDirName & "_" & EDIT_FILE_PREFIX     & sAddDate
            sDstOrgFilePath    = sDstParEvaDirPath & "\" & sSrcDirName & "_" & ORIGINAL_FILE_PREFIX & sAddDate
            sDstShrtctFilePath = sDstParEvaDirPath & "\" & sSrcDirName & "_" & SHORTCUT_FILE_SUFFIX & sAddDate & ".lnk"
            
            '�t�H���_�R�s�[
            objFSO.CopyFolder sSelectedPath, sDstCpyFilePath, True
            objFSO.CopyFolder sSelectedPath, sDstOrgFilePath, True
            
            '�V���[�g�J�b�g�쐬
            With objWshShell.CreateShortcut( sDstShrtctFilePath )
                .TargetPath = sOrgDirPath
                .Save
            End With
            
        '### �t�@�C��/�t�H���_�ȊO ###
        Else
            MsgBox "�I�����ꂽ�I�u�W�F�N�g�����݂��܂���" & vbNewLine & vbNewLine & sSelectedPath, vbOKOnly, PROG_NAME
            MsgBox "�����𒆒f���܂��B", vbOKOnly, PROG_NAME
            bIsContinue = False
        End If
    Next
    
    Set objFSO = Nothing
    Set objWshShell = Nothing
    
    MsgBox "�V���[�g�J�b�g���R�s�[�t�@�C���̍쐬���������܂����I", vbOKOnly, PROG_NAME
Else
    'Do Nothing
End If

' ==================================================================
' = �T�v    ������������t�@�C��/�t�H���_���ɓK�p�ł���`���ɕϊ�����
' = ����    sDateRaw    String  [in]    �����i��F2017/8/5 12:59:58�j
' = �ߒl                String          �����i��F20170805-125958�j
' = �o��    �Ȃ�
' ==================================================================
Public Function ConvDate2String( _
    ByVal sDateRaw _
)
    Dim sSearchPattern
    Dim sTargetStr
    sSearchPattern = "(\w{4})/(\w{1,2})/(\w{1,2}) (\w{1,2}):(\w{1,2}):(\w{1,2})"
    sTargetStr = sDateRaw
    
    Dim oRegExp
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = sSearchPattern                '�����p�^�[����ݒ�
    oRegExp.IgnoreCase = True                       '�啶���Ə���������ʂ��Ȃ�
    oRegExp.Global = True                           '������S�̂�����
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)  '�p�^�[���}�b�`���s
    Dim sDateStr
    With oMatchResult(0)
        sDateStr = Right( .SubMatches(0), 2 ) & _
                   String( 2 - Len( .SubMatches(1) ), "0" ) & .SubMatches(1) & _
                   String( 2 - Len( .SubMatches(2) ), "0" ) & .SubMatches(2) & _
                   String( 2 - Len( .SubMatches(3) ), "0" ) & .SubMatches(3) & _
                   String( 2 - Len( .SubMatches(4) ), "0" ) & .SubMatches(4)
    End With
    Set oMatchResult = Nothing
    Set oRegExp = Nothing
    ConvDate2String = sDateStr
End Function

' ==================================================================
' = �T�v    �t�H���_�I���_�C�A���O��\������
' = ����    sInitPath   String  [in]  �f�t�H���g�t�H���_�p�X
' = �ߒl                String        �t�H���_�I������
' = �o��    �E���݂��Ȃ��t�H���_�p�X��I�������ꍇ�A�󕶎����ԋp����
' =         �E�L�����Z�������������ꍇ�A�󕶎����ԋp����
' ==================================================================
Private Function ShowFolderSelectDialog( _
    ByVal sInitPath _
)
    Const msoFileDialogFolderPicker = 4
    Const xlMinimized = -4140
    
    Dim objExcel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False '��\���ɂ��Ă�����ۂɂ�����ƕ\�����ꂿ�Ⴄ�B
    objExcel.WindowState = xlMinimized '��L���R����ŏ��������Ƃ��B
    
    Dim fdDialog
    Set fdDialog = objExcel.FileDialog(msoFileDialogFolderPicker)
    fdDialog.Title = "�t�H���_��I�����Ă��������i�󗓂̏ꍇ�͐e�t�H���_���I������܂��j"
    If sInitPath = "" Then
        'Do Nothing
    Else
        If Right(sInitPath, 1) = "\" Then
            fdDialog.InitialFileName = sInitPath
        Else
            fdDialog.InitialFileName = sInitPath & "\"
        End If
    End If
    
    '�_�C�A���O�\��
    Dim lResult
    lResult = fdDialog.Show()
    If lResult <> -1 Then '�L�����Z������
        ShowFolderSelectDialog = ""
    Else
        Dim sSelectedPath
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FolderExists( sSelectedPath ) Then
            ShowFolderSelectDialog = sSelectedPath
        Else
            ShowFolderSelectDialog = ""
        End If
    End If
    
    Set fdDialog = Nothing
End Function
'   Call Test_ShowFolderSelectDialog()
    Private Sub Test_ShowFolderSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        
        Dim sInitPath
        sInitPath = objWshShell.SpecialFolders("Desktop")
        'sInitPath = ""
        
        MsgBox ShowFolderSelectDialog( sInitPath )
    End Sub

Public Function SetFileAttributes( _
    ByVal sFilePath, _
    ByVal sDateRaw _
)
    Const SET_ATTR_READONLY = 1     ' �ǂݎ���p�t�@�C��
    Const SET_ATTR_HIDDEN   = 2     ' �B���t�@�C��
    Const SET_ATTR_SYSTEM   = 4     ' �V�X�e���E�t�@�C��
    Const SET_ATTR_ARCHIVE  = 32    ' �O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1
    
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile("test.txt ")
End Function

