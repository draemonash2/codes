Option Explicit

'===============================================================================
'�y�T�v�z
'   xEV�V�~�����[�^���o�͂����������OCSV�𐮌`���ACANape���O�g�`�o�͗p�̌`���ɕϊ�����B
'       �E�uDatatype�v���t�^
'           �iDataType�� data_type_list.csv ���擾�j
'       �E�ϐ��V���{��������z�񎯕ʎq������ (�f�t�H���g�ŋ@�\OFF)
'           ex) ram[0]:1 �� ram_0:1
'
'�y�g�p���@�z
'   �g�p���@�͂Q�ʂ�B
'       ���t�H���_�z���̑S�������O(CSV)�𐮌`�������ꍇ
'           �P�D�udata_type_list.csv�v���쐬�B
'                 ex) AAA:1[1],uint8
'                     AAA:1[2],uint8
'                     BBB:1,sint16
'                     CCC:2,double
'           �Q�D���`�Ώۂ̎������O(CSV)�Ɠ��K�w�ȏ�̃t�H���_��
'               �u�������OCSV���`�c�[��.vbs�v�Ɓudata_type_list.csv�v���i�[�B
'           �R�D�u�������OCSV���`�c�[��.vbs�v�����s�B
'       ���P�t�@�C���̂ݐ��`�������ꍇ
'           �P�D�udata_type_list.csv�v���쐬�B
'           �Q�D�u�������OCSV���`�c�[��.vbs�v�Ɠ��K�w�̃t�H���_��
'               �udata_type_list.csv�v���i�[�B
'           �R. ���`�������������O(CSV)���u�������OCSV���`�c�[��.vbs�v��drag&drop����B
'
'�y�ڍ׎d�l�z
'   �E�t�@�C���̐擪��"TimeStamp"�ƋL�ڂ��ꂽ.csv�t�@�C�����������O(CSV)�Ɖ��߂���B
'   �E�ȉ��̂悤�Ȓǉ��ݒ肪�\�B
'     - �ϐ��V���{��������z�񎯕ʎq����������@�\�̗L������
'         �� REPLACE_RAM_SYMBOL = True:�L�� / False:����
'     - �������O(CSV)�̃o�b�N�A�b�v���쐬�L��
'         �� CREATE_BACKUP_FILE = True:�o�b�N�A�b�v�t�@�C���쐬 / False:�㏑��
'     - ���`�������̃��b�Z�[�W�o�͗L��
'         �� OUTPUT_FINISH_MESSAGE = True:�o�� / False:�o�͂��Ȃ�
'   �Edata_type_list.csv �ɂ���
'     - data_type_list.csv �����݂��Ȃ��ꍇ�́A���ׂ� uint8 �Ɖ��߂���B
'     - data_type_list.csv �ɑ��݂��Ȃ�RAM�́Auint8 �Ɖ��߂���B
'
'�y���������z
'   1.0.0   2019/07/01  ����    �E�V�K�쐬
'===============================================================================

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )          'GetFileList3()
                                                                     'GetFileNotExistPath()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\Collection.vbs" )          'ReadTxtFileToCollection()
                                                                     'WriteTxtFileFrCollection()
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\ProgressBarCscript.vbs" )  'Class ProgressBar

'===============================================================================
' �ݒ�
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"
CONST DEFAULT_DATA_TYPE = "uint8"
CONST CREATE_BACKUP_FILE = False
CONST REPLACE_RAM_SYMBOL = False
CONST OUTPUT_FINISH_MESSAGE = True

'===============================================================================
' �{����
'===============================================================================
Const RAMNAME_ROW_KEYWORD = "TimeStamp,"
Const DATATYPE_ROW_KEYWORD = "DataType"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim objPrgrsBar
Set objPrgrsBar = New ProgressBar
objPrgrsBar.Message = "�������OCSV���`��..."

'*****************************
' �������OCSV�t�@�C�����X�g�擾
'*****************************
dim cCsvFileList
Set cCsvFileList = CreateObject("System.Collections.ArrayList")

Dim sRootDirPath
sRootDirPath = objFSO.GetParentFolderName( WScript.ScriptFullName )

If WScript.Arguments.Count = 0 Then
    dim cFileList
    Set cFileList = CreateObject("System.Collections.ArrayList")
    call GetFileList3(sRootDirPath, cFileList, 1)
    
    dim sFilePath
    for each sFilePath in cFileList
        if objFSO.GetExtensionName(sFilePath) = "csv" And _
           objFSO.GetFileName(sFilePath) <> DATA_TYPE_LIST_FILE_NAME then
            cCsvFileList.add sFilePath
        end if
    next
    Set cFileList = Nothing
ElseIf WScript.Arguments.Count = 1 And _
    objFSO.FileExists(WScript.Arguments(0)) Then
    cCsvFileList.add WScript.Arguments(0)
Else
    WScript.Echo "�w�肷������̐�������Ă��܂�:" & WScript.Arguments.Count
    WScript.Quit
End If

'*****************************
' DataType�ꗗ�擾
'*****************************
dim dDataTypeList
Set dDataTypeList = CreateObject("Scripting.Dictionary")

Dim sDataTypeListFilePath
sDataTypeListFilePath = sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME

dim objTxtFile
If objFSO.FileExists(sDataTypeListFilePath) Then
    set objTxtFile = objFSO.OpenTextFile(sDataTypeListFilePath, 1)

    dim objWords
    Dim sTxtLine
    Do Until objTxtFile.AtEndOfStream
        sTxtLine = objTxtFile.ReadLine
        objWords = split(sTxtLine, ",")
        if REPLACE_RAM_SYMBOL = True then
            objWords(0) = RenameRamSymbol(objWords(0))
        end if
        On Error Resume Next '�d���L�[���������疳��
        dDataTypeList.Add objWords(0), objWords(1) 'RamName DataType
        On Error Goto 0
    Loop
    objTxtFile.Close
Else
    'Do Nothing
End If

'*****************************
' �������OCSV���`
'*****************************
dim sCsvFilePath
Dim lProcIdx
Dim lProcNum
lProcIdx = 0
lProcNum = cCsvFileList.Count
Call objPrgrsBar.Update(lProcIdx, lProcNum)
for each sCsvFilePath In cCsvFileList
    
    '*** �������OCSV�I�[�v�� ***
    dim cFileContents
    Set cFileContents = CreateObject("System.Collections.ArrayList")
    call ReadTxtFileToCollection(sCsvFilePath, cFileContents)
    
    '*** �������O�t�@�C���`�F�b�N ***
    If Left(cFileContents(0), len(RAMNAME_ROW_KEYWORD)) = RAMNAME_ROW_KEYWORD Then
        
        '*** �o�b�N�A�b�v�o�� ***
        If CREATE_BACKUP_FILE = True then
            Dim sCsvBakFilePathRaw
            Dim sCsvBakFilePath
            Dim lBakFileIdx
            sCsvBakFilePathRaw = sCsvFilePath & ".bak"
            sCsvBakFilePath = sCsvBakFilePathRaw
            lBakFileIdx = 1
            Do While objFSO.FileExists( sCsvBakFilePath )
                sCsvBakFilePath = sCsvBakFilePathRaw & lBakFileIdx
                lBakFileIdx = lBakFileIdx + 1
            Loop
            objFSO.CopyFile sCsvFilePath, sCsvBakFilePath
        End If
        
        '*** �ϐ����u�� ***
        if REPLACE_RAM_SYMBOL = True then
            cFileContents(0) = RenameRamSymbol(cFileContents(0))
        end if
        
        '*** Datatype�u��or�}�� ***
        Dim vRamNames
        vRamNames = Split(cFileContents(0), ",")
        Dim sRamName
        Dim sDataTypeLine
        Dim lIdx
        lIdx = 0
        for each sRamName In vRamNames
            If lIdx = 0 Then
                sDataTypeLine = DATATYPE_ROW_KEYWORD
            else
                '���łɒu���ς�
                'if REPLACE_RAM_SYMBOL = True then
                '   sRamName = RenameRamSymbol(sRamName)
                'end if
                if dDataTypeList.Exists(sRamName) then
                    sDataTypeLine = sDataTypeLine & "," & dDataTypeList.Item(sRamName)
                else
                    sDataTypeLine = sDataTypeLine & "," & DEFAULT_DATA_TYPE
                end if
            end if
            lIdx = lIdx + 1
        next
        If Left(cFileContents(1), len(DATATYPE_ROW_KEYWORD)) = DATATYPE_ROW_KEYWORD Then
            cFileContents(1) = sDataTypeLine
        Else
            cFileContents.Insert 1, sDataTypeLine
        End If
        
        '*** CSV�o�� ***
        call WriteTxtFileFrCollection(sCsvFilePath, cFileContents, True)
    Else
        'Do Nothing
    End If
    
    lProcIdx = lProcIdx + 1
    Call objPrgrsBar.Update(lProcIdx, lProcNum)
    
    Set cFileContents = Nothing
next

Set objFSO = Nothing
Set cCsvFileList = Nothing
Set dDataTypeList = Nothing

IF OUTPUT_FINISH_MESSAGE = True Then
    MsgBox "�������OCSV ���`����!"
End If

'===============================================================================
' �֐�
'===============================================================================
Private Function RenameRamSymbol( _
    byval sTrgtWord _
)
    Dim sOutWord
    sOutWord = sTrgtWord
    sOutWord = Replace(sOutWord, "[", "_")
    sOutWord = Replace(sOutWord, "]", "")
    RenameRamSymbol = sOutWord
End Function

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

