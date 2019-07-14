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
'         �� REPLACE_RAM_NAME = True:�L�� / False:����
'     - ���`�������̃��b�Z�[�W�o�͗L��
'         �� OUTPUT_FINISH_MESSAGE = True:�o�� / False:�o�͂��Ȃ�
'
'�y���������z
'   1.0.0   2019/07/01  ����    �E�V�K�쐬
'===============================================================================
'===============================================================================
' �ݒ�
'===============================================================================
CONST DATA_TYPE_LIST_FILE_NAME = "data_type_list.csv"
CONST REPLACE_RAM_NAME = False
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
                if REPLACE_RAM_NAME = True then
                    sRamName = ReplaceKeyword(sRamName)
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
Private Function ReplaceKeyword( _
    byval sTrgtWord _
)
    Dim sOutWord
    sOutWord = sTrgtWord
    sOutWord = Replace(sOutWord, "[", "_")
    sOutWord = Replace(sOutWord, "]", "")
    ReplaceKeyword = sOutWord
End Function

' ==================================================================
' = �T�v    �t�@�C��/�t�H���_�p�X�ꗗ���擾����
' = ����    sTrgtDir        String      [in]    �Ώۃt�H���_
' = ����    cFileList       Collections [out]   �t�@�C��/�t�H���_�p�X�ꗗ
' = ����    lFileListType   Long        [in]    �擾����ꗗ�̌`��
' =                                                 0�F����
' =                                                 1:�t�@�C��
' =                                                 2:�t�H���_
' =                                                 ����ȊO�F�i�[���Ȃ�
' = �ߒl    �Ȃ�
' = �o��    �EDir �R�}���h�ɂ��t�@�C���ꗗ�擾�BGetFileList() ���������B
' =         �EArray�R���N�V�����Ɋi�[����
' ==================================================================
Public Function GetFileList3( _
    ByVal sTrgtDir, _
    ByRef cFileList, _
    ByVal lFileListType _
)
    Dim objFSO  'FileSystemObject�̊i�[��
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    
    'Dir �R�}���h���s�i�o�͌��ʂ��ꎞ�t�@�C���Ɋi�[�j
    Dim sTmpFilePath
    Dim sExecCmd
    sTmpFilePath = WScript.CreateObject( "WScript.Shell" ).CurrentDirectory & "\Dir.tmp"
    Select Case lFileListType
        Case 0:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a > """ & sTmpFilePath & """"
        Case 1:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:a-d > """ & sTmpFilePath & """"
        Case 2:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:d > """ & sTmpFilePath & """"
        Case Else: sExecCmd = ""
    End Select
    With CreateObject("Wscript.Shell")
        .Run "cmd /c" & sExecCmd, 7, True
    End With
    
    Dim objFile
    Dim sTextAll
    On Error Resume Next
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile( sTmpFilePath, 1 )
        If Err.Number = 0 Then
            sTextAll = objFile.ReadAll
            sTextAll = Left( sTextAll, Len( sTextAll ) - Len( vbNewLine ) ) '�����ɉ��s���t�^����Ă��܂����߁A�폜
            Dim vFileList
            vFileList = Split( sTextAll, vbNewLine )
            Dim sFilePath
            For Each sFilePath In vFileList
                cFileList.add sFilePath
            Next
            objFile.Close
        Else
            WScript.Echo "�t�@�C�����J���܂���: " & Err.Description
        End If
        Set objFile = Nothing   '�I�u�W�F�N�g�̔j��
    Else
        WScript.Echo "�G���[ " & Err.Description
    End If  
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    '�I�u�W�F�N�g�̔j��
    On Error Goto 0
End Function
'   Call Test_GetFileList3()
    Private Sub Test_GetFileList3()
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Dim sCurDir
        sCurDir = objFSO.GetParentFolderName( WScript.ScriptFullName )
        
        msgbox sCurDir
        
        Dim cFileList
        Set cFileList = CreateObject("System.Collections.ArrayList")
        Call GetFileList3( sCurDir, cFileList, 1 )
        
        dim sFilePath
        dim sOutput
        sOutput = ""
        for each sFilePath in cFileList
            sOutput = sOutput & vbNewLine & sFilePath
        next
        MsgBox sOutput
    End Sub

' ==================================================================
' = �T�v    �w��t�@�C���p�X�����݂���ꍇ�A"_XXX" ��t�^���ĕԋp����
' = ����    sTrgtFilePath   String      [in]    �Ώۃp�X
' = �ߒl                    String              �t�^��p�X
' = �o��    �{�֐��ł́A�t�@�C���͍쐬���Ȃ��B
' = �ˑ�lib �Ȃ�
' ==================================================================
Public Function GetFileNotExistPath( _
    ByVal sTrgtFilePath _
)
    Dim lIdx
    Dim objFSO
    Dim sFileParDirPath
    Dim sFileBaseName
    Dim sFileExtName
    Dim sCreFilePath
    Dim bIsTrgtPathExists
    
    lIdx = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sCreFilePath = sTrgtFilePath
    bIsTrgtPathExists = False
    Do While objFSO.FileExists( sCreFilePath )
        bIsTrgtPathExists = True
        lIdx = lIdx + 1
        sFileParDirPath = objFSO.GetParentFolderName( sTrgtFilePath )
        sFileBaseName = objFSO.GetBaseName( sTrgtFilePath ) & "_" & String( 3 - len(lIdx), "0" ) & lIdx
        sFileExtName = objFSO.GetExtensionName( sTrgtFilePath )
        If sFileExtName = "" Then
            sCreFilePath = sFileParDirPath & "\" & sFileBaseName
        Else
            sCreFilePath = sFileParDirPath & "\" & sFileBaseName & "." & sFileExtName
        End If
    Loop
    GetFileNotExistPath = sCreFilePath
End Function
'   Call Test_GetFileNotExistPath()
    Private Sub Test_GetFileNotExistPath()
        Dim sOutStr
        sOutStr = ""
        sOutStr = sOutStr & vbNewLine & "*** test start! ***"
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\MacroBook\lib\FileSys.bas")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\MacroBook\lib\FileSy.bas")
        sOutStr = sOutStr & vbNewLine & GetFileNotExistPath("C:\codes\vba\AddIns\UserDefFuncs.bas")
        sOutStr = sOutStr & vbNewLine & "*** test finished! ***"
        MsgBox sOutStr
    End Sub

' ==================================================================
' = �T�v    �e�L�X�g�t�@�C���̒��g��z��Ɋi�[
' = ����    sTrgtFilePath   String      [in]    �t�@�C���p�X
' = ����    cFileContents   Collections [out]   �t�@�C���̒��g
' = �ߒl    �ǂݏo������    Boolean             �ǂݏo������
' =                                                 True:�t�@�C������
' =                                                 False:����ȊO
' = �o��    �Ȃ�
' ==================================================================
Public Function ReadTxtFileToCollection( _
    ByVal sTrgtFilePath, _
    ByRef cFileContents _
)
    On Error Resume Next
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(sTrgtFilePath) Then
        Dim objTxtFile
        Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 1, True)
        
        If Err.Number = 0 Then
            Do Until objTxtFile.AtEndOfStream
                cFileContents.add objTxtFile.ReadLine
            Loop
            ReadTxtFileToCollection = True
        Else
            ReadTxtFileToCollection = False
        '   WScript.Echo "�G���[ " & Err.Description
        End If
        
        objTxtFile.Close
    Else
        ReadTxtFileToCollection = False
    End If
    On Error Goto 0
End Function
'   Call Test_OpenTxtFile2Array()
    Private Sub Test_OpenTxtFile2Array()
        Dim cFileList
        Set cFileList = CreateObject("System.Collections.ArrayList")
        sFilePath = "C:\codes\vbs\��������CSV���`�c�[��\data_type_list_.csv"
        Dim bRet
        bRet = ReadTxtFileToCollection( sFilePath, cFileList )
        
        dim sFilePath
        dim sOutput
        sOutput = ""
        for each sFilePath in cFileList
            sOutput = sOutput & vbNewLine & sFilePath
        next
        MsgBox bRet
        MsgBox sOutput
    End Sub

' ==================================================================
' = �T�v    �z��̒��g���e�L�X�g�t�@�C���ɏ����o��
' = ����    sTrgtFilePath   String      [in]    �t�@�C���p�X
' = ����    cFileContents   Collections [in]    �t�@�C���̒��g
' = ����    bOverwrite      Boolean     [in]    True:�㏑���AFalse:�V�K�t�@�C��
' = �ߒl    �����o������    Boolean             �����o������
' =                                                 True:�����o������
' =                                                 False:����ȊO
' = �o��    �Ȃ�
' = �ˑ�lib FileSystem.vbs/GetFileNotExistPath
' ==================================================================
Public Function WriteTxtFileFrCollection( _
    ByVal sTrgtFilePath, _
    ByRef cFileContents, _
    ByVal bOverwrite _
)
    On Error Resume Next
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objTxtFile
    If bOverwrite = True Then
        'Do Nothing
    Else
        Dim sInTrgtFilePath
        sInTrgtFilePath = sTrgtFilePath
        sTrgtFilePath = GetFileNotExistPath(sInTrgtFilePath)
    End If
    Set objTxtFile = objFSO.OpenTextFile(sTrgtFilePath, 2, True)
    
    If Err.Number = 0 Then
        Dim sFileLine
        For Each sFileLine In cFileContents
            objTxtFile.WriteLine sFileLine
        Next
        WriteTxtFileFrCollection = True
    Else
        WriteTxtFileFrCollection = False
    '   WScript.Echo "�G���[ " & Err.Description
    End If
    
    objTxtFile.Close
    On Error Goto 0
End Function
'   Call Test_WriteTxtFileFrCollection()
    Private Sub Test_WriteTxtFileFrCollection()
        Dim cFileContents
        Set cFileContents = CreateObject("System.Collections.ArrayList")
        cFileContents.Add "a"
        cFileContents.Add "b"
        cFileContents.Add "d"
        cFileContents.Add "e"
        cFileContents.Insert 1, "c"
        DIm sTrgtFilePath
        sTrgtFilePath = "C:\codes\vbs\_lib\Test.csv.bak"
        call WriteTxtFileFrCollection( sTrgtFilePath, cFileContents, False )
    End Sub

' progrress bar cscript class v1.00
Class ProgressBar
    Private sStatus
    Private objFSO
    Private objWshShell
    
    Private Sub Class_Initialize
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objWshShell = WScript.CreateObject("WScript.Shell")
        Dim sExeFileName
        sExeFileName = LCase(objFSO.GetFileName(WScript.FullName))
        if sExeFileName = "cscript.exe" then
            'Do Nothing
        else
            objWshShell.Run "cscript //nologo """ & Wscript.ScriptFullName & """", 1, False
            Wscript.Quit
        end if
    End Sub
    
    Private Sub Class_Terminate
        Set objFSO = Nothing
        Set objWshShell = Nothing
    End Sub
    
    ' ==================================================================
    ' = �T�v    ���b�Z�[�W���X�V����
    ' = ����    sProgMsg      String   [in] ���b�Z�[�W
    ' = �ߒl    �Ȃ�
    ' = �o��    �Ȃ�
    ' ==================================================================
    Public Property Let Message( _
        ByVal sMessage _
    )
        if sStatus = "Update" then
            Wscript.StdOut.Write vbCrLf
        end if
        Wscript.StdOut.Write sMessage & vbCrLf
        sStatus = "Message"
    End Property
    
    ' ==================================================================
    ' = �T�v    �i�����X�V����
    ' = ����    lBunsi      Long   [in] �i��
    ' = ����    lBunbo      Long   [in] �i���ő�l
    ' = �ߒl    �Ȃ�
    ' = �o��    �Ȃ�
    ' ==================================================================
    Public Sub Update( _
        ByVal lBunsi, _
        ByVal lBunbo _
    )
        '�p�[�Z���e�[�W�v�Z
        Dim iPercentage
        Dim sPercentage
        iPercentage = Cint((lBunsi / lBunbo) * 100)
        sPercentage = iPercentage & "%"
        sPercentage = String(4 - Len(sPercentage), " ") & sPercentage
        
        '�i���o�[
        Dim sProgressBar
        sProgressBar = String(Cint(iPercentage/5), "=") & ">" & String(20 - Cint(iPercentage/5), " ")
        
        '�`��
        Wscript.StdOut.Write sPercentage & " |" & sProgressBar & "| " & lBunsi & "/" & lBunbo & vbCr
        sStatus = "Update"
    End Sub
    
    ' ==================================================================
    ' = �T�v    �v���O���X�o�[���I������
    ' = ����    �Ȃ�
    ' = �ߒl    �Ȃ�
    ' = �o��    cscript�͏I���ł��Ȃ�
    ' ==================================================================
'   Public Function Quit()
'       gobjExplorer.Document.Body.Style.Cursor = "default"
'       gobjExplorer.Quit
'   End Function
    
End Class
    If WScript.ScriptName = "ProgressBarCscript.vbs" Then
        Call Test_ProgressBar
    End If
    Private Sub Test_ProgressBar
        Dim lProcIdx
        Dim lProcNum
        Dim objPrgrsBar
        Set objPrgrsBar = New ProgressBar
        
        '#�����P
        objPrgrsBar.Message = "�������� ���s!"
        lProcNum = 255
        For lProcIdx = 1 To lProcNum
            WScript.Sleep 1
            Call objPrgrsBar.Update(lProcIdx, lProcNum)
        Next
        
        '#�����Q
        objPrgrsBar.Message = "�Z������ ���s!"
        lProcNum= 10
        For lProcIdx = 1 To lProcNum
            WScript.Sleep 45
            Call objPrgrsBar.Update(lProcIdx, lProcNum)
        Next
        
        objPrgrsBar.Message = "Complete!!"
        msgbox "�I�����܂���"
    End Sub
