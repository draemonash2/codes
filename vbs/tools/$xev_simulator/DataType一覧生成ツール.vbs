Option Explicit

'===============================================================================
'�y�T�v�z
'   �������OCSV����ϐ��V���{������DataType�𒊏o���āudata_type_list.csv�v�𐶐�����
'
'�y�g�p���@�z
'   �g�p���@�͂Q�ʂ�B
'       ���t�H���_�z���̑S�������O(CSV)����DataType�𒊏o�������ꍇ
'           �P�D���o�Ώۂ̎������O(CSV)�Ɠ��K�w�ȏ�̃t�H���_��
'               �uDataType�ꗗ�����c�[��.vbs�v���i�[�B
'           �Q�D�uDataType�ꗗ�����c�[��.vbs�v�����s����B
'       ���P�t�@�C������̂ݒ��o�������ꍇ
'           �P�D���o�������������O(CSV)���uDataType�ꗗ�����c�[��.vbs�v��drag&drop����B
'
'�y�ڍ׎d�l�z
'   �E�t�@�C���̐擪��"TimeStamp"�ƋL�ڂ��ꂽ.csv�t�@�C�����������O(CSV)�Ɖ��߂���B
'   �E�ȉ��̂悤�Ȓǉ��ݒ肪�\�B
'     - �ϐ��V���{��������z�񎯕ʎq����������@�\�̗L������
'         �� REPLACE_RAM_SYMBOL = True:�L�� / False:����
'     - ���`�������̃��b�Z�[�W�o�͗L��
'         �� OUTPUT_FINISH_MESSAGE = True:�o�� / False:�o�͂��Ȃ�
'
'�y���������z
'   1.0.0   2019/07/01  ����    �E�V�K�쐬
'===============================================================================

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )          'GetFileList3()
                                                            'GetFileNotExistPath()
Call Include( "C:\codes\vbs\_lib\Collection.vbs" )          'ReadTxtFileToCollection()
                                                            'WriteTxtFileFrCollection()
Call Include( "C:\codes\vbs\_lib\ProgressBarCscript.vbs" )  'Class ProgressBar

'===============================================================================
' �ݒ�
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"
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
objPrgrsBar.Message = "DataType�ꗗ������..."

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
' DataType�擾
'*****************************
dim cDataTypeList
Set cDataTypeList = CreateObject("System.Collections.ArrayList")
dim dDataTypeListDupChk '�d���`�F�b�N�p
set dDataTypeListDupChk = CreateObject("Scripting.Dictionary")
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
    If Left(cFileContents(0), len(RAMNAME_ROW_KEYWORD)) = RAMNAME_ROW_KEYWORD And _
       Left(cFileContents(1), len(DATATYPE_ROW_KEYWORD)) = DATATYPE_ROW_KEYWORD Then
        
        '*** DataType�擾 ***
        Dim vRamNames
        Dim vDataTypes
        vRamNames = Split(cFileContents(0), ",")
        vDataTypes = Split(cFileContents(1), ",")
        Dim sRamName
        Dim lIdx
        lIdx = 0
        for each sRamName In vRamNames
            If lIdx = 0 Then '1��ڂ͖���
                'Do Nothing
            else
                if REPLACE_RAM_SYMBOL = True then
                    sRamName = RenameRamSymbol(sRamName)
                end if
                Dim sDataTypeListLine
                sDataTypeListLine = sRamName & "," & vDataTypes(lIdx)
                If Not dDataTypeListDupChk.Exists( sDataTypeListLine ) Then
                    cDataTypeList.Add sDataTypeListLine
                    dDataTypeListDupChk.Add sDataTypeListLine, ""
                end if
            end if
            lIdx = lIdx + 1
        next
    Else
        'Do Nothing
    End If
    
    lProcIdx = lProcIdx + 1
    Call objPrgrsBar.Update(lProcIdx, lProcNum)
    
    Set cFileContents = Nothing
next

'*****************************
' DataType�ꗗ�o��
'*****************************
call WriteTxtFileFrCollection(sRootDirPath & "\" & DATA_TYPE_LIST_FILE_NAME, cDataTypeList, True)

Set objFSO = Nothing
Set cCsvFileList = Nothing
Set cDataTypeList = Nothing
set dDataTypeListDupChk = Nothing

IF OUTPUT_FINISH_MESSAGE = True Then
    MsgBox "DataType�ꗗ ��������!"
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
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

