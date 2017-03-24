Attribute VB_Name = "FileSys"
Option Explicit

' file system library v1.2

Public Enum E_PATH_TYPE
    PATH_TYPE_FILE
    PATH_TYPE_DIRECTORY
End Enum
 
Public Type T_PATH_LIST
    sPath As String
    sName As String
    ePathType As E_PATH_TYPE
End Type
 
Public Enum T_SYSOBJ_TYPE
    SYSOBJ_NOT_EXIST
    SYSOBJ_FILE
    SYSOBJ_DIRECTORY
End Enum
  
'�Q�Ɛݒ�uMicrosoft ActiveX Data Objects 6.1 Liblary�v���`�F�b�N���邱�ƁI
' ============================================
' = �T�v    �t�@�C���̓��e��z��ɓǂݍ��ށB
' = ����    sFilePath   String   ���͂���t�@�C���p�X
' =         sCharSet    String   �L�����N�^�Z�b�g
' = �ߒl                String() �t�@�C�����e
' = �o��    �Ȃ�
' ============================================
Public Function InputTxtFile( _
    ByRef sFilePath As String, _
    Optional ByVal sCharSet As String _
) As String()
    Dim lLineCnt As Long: lLineCnt = 0
    Dim asRetStr() As String
    Dim oTxtObj As Object
    
    Set oTxtObj = CreateObject("ADODB.Stream")
    
    With oTxtObj
        .Type = adTypeText           '�I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ𕶎���^�Ɏw�肷��
        .Charset = sCharSet
        .Open
        .LoadFromFile (sFilePath)
        
        lLineCnt = 0
        Do While Not .EOS
            ReDim Preserve asRetStr(lLineCnt)
            asRetStr(lLineCnt) = .ReadText(adReadLine)
            lLineCnt = lLineCnt + 1
        Loop
        
        .Close
    End With
    
    Set oTxtObj = Nothing
    
    InputTxtFile = asRetStr
    
End Function

' ============================================
' = �T�v    �z��̓��e���t�@�C���ɏ������ށB
' = ����    sFilePath     String  [in]  �o�͂���t�@�C���p�X
' =         asFileLine()  String  [in]  �o�͂���t�@�C���̓��e
' = �ߒl    �Ȃ�
' = �o��    �Ȃ�
' ============================================
Public Function OutputTxtFile( _
    ByVal sFilePath As String, _
    ByRef asFileLine() As String, _
    Optional ByVal sCharSet As String _
)
    Dim oTxtObj As Object
    Dim lLineIdx As Long
    
    If Sgn(asFileLine) = 0 Then
        'Do Nothing
    Else
        Set oTxtObj = CreateObject("ADODB.Stream")
        With oTxtObj
            .Type = adTypeText
            .Charset = sCharSet
            .Open
            
            '�z���1�s���I�u�W�F�N�g�ɏ�������
            For lLineIdx = 0 To UBound(asFileLine)
                .WriteText asFileLine(lLineIdx), adWriteLine
            Next lLineIdx
            
            .SaveToFile (sFilePath), adSaveCreateOverWrite    '�I�u�W�F�N�g�̓��e���t�@�C���ɕۑ�
            .Close
        End With
    End If
    
    Set oTxtObj = Nothing
End Function

'�t�H���_�����ɑ��݂��Ă���ꍇ�͉������Ȃ�
Public Function CreateDirectry( _
    ByVal sDirPath As String _
)
    Dim sParentDir As String
    Dim oFileSys As Object
 
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
 
    sParentDir = oFileSys.GetParentFolderName(sDirPath)
 
    '�e�f�B���N�g�������݂��Ȃ��ꍇ�A�ċA�Ăяo��
    If oFileSys.FolderExists(sParentDir) = False Then
        Call CreateDirectry(sParentDir)
    End If
 
    '�f�B���N�g���쐬
    If oFileSys.FolderExists(sDirPath) = False Then
        oFileSys.CreateFolder sDirPath
    End If
 
    Set oFileSys = Nothing
End Function

'atPathList() �Ƀt�@�C�����X�g���i�[�����B
Public Function GetFileList( _
    ByVal sTargetDir As String, _
    ByRef atPathList() As T_PATH_LIST _
)
    Dim oFolder As Object
    Dim oSubFolder As Object
    Dim oFile As Object
    Dim lLastIdx As Long
 
    Set oFolder = CreateObject("Scripting.FileSystemObject").GetFolder(sTargetDir)
 
    '*** �t�H���_�� ***
    If Sgn(atPathList) = 0 Then
        ReDim Preserve atPathList(0)
    Else
        ReDim Preserve atPathList(UBound(atPathList) + 1)
    End If
    lLastIdx = UBound(atPathList)
    atPathList(lLastIdx).sPath = oFolder.Path
    atPathList(lLastIdx).sName = oFolder.Name
    atPathList(lLastIdx).ePathType = PATH_TYPE_DIRECTORY
 
    '�t�H���_���̃T�u�t�H���_���
    '�i�T�u�t�H���_���Ȃ���΃��[�v���͒ʂ�Ȃ��j
    For Each oSubFolder In oFolder.SubFolders
        Call GetFileList(oSubFolder.Path, atPathList) '�ċA�I�Ăяo��
    Next oSubFolder
 
    '*** �t�@�C���� ***
    For Each oFile In oFolder.Files
        If Sgn(atPathList) = 0 Then
            ReDim Preserve atPathList(0)
        Else
            ReDim Preserve atPathList(UBound(atPathList) + 1)
        End If
        lLastIdx = UBound(atPathList)
        atPathList(lLastIdx).sPath = oFile.Path
        atPathList(lLastIdx).sName = oFile.Name
        atPathList(lLastIdx).ePathType = PATH_TYPE_FILE
    Next oFile
End Function

' ==================================================================
' = �T�v    �t�@�C��/�t�H���_�p�X�ꗗ���擾����
' = ����    sTrgtDir        String      [in]    �Ώۃt�H���_
' = ����    asFileList      Variant     [out]   �t�@�C��/�t�H���_�p�X�ꗗ
' = ����    lFileListType   Long        [in]    �擾����ꗗ�̌`��
' =                                                 0�F����
' =                                                 1:�t�@�C��
' =                                                 2:�t�H���_
' =                                                 ����ȊO�F�i�[���Ȃ�
' = �ߒl    �Ȃ�
' = �o��    �EDir �R�}���h�ɂ��t�@�C���ꗗ�擾�BGetFileList() ���������B
' =         �EasFileList �͔z��^�ł͂Ȃ��o���A���g�^�Ƃ��Ē�`����
' =           �K�v�����邱�Ƃɒ��ӁI
' ==================================================================
Public Function GetFileList2( _
    ByVal sTrgtDir, _
    ByRef asFileList, _
    ByVal lFileListType _
)
    Dim objFSO As Object 'FileSystemObject�̊i�[��
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Dir �R�}���h���s�i�o�͌��ʂ��ꎞ�t�@�C���Ɋi�[�j
    Dim sTmpFilePath As String
    Dim sExecCmd As String
    sTmpFilePath = CreateObject("WScript.Shell").CurrentDirectory & "\Dir.tmp"
    Select Case lFileListType
        Case 0:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a > """ & sTmpFilePath & """"
        Case 1:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:a-d > """ & sTmpFilePath & """"
        Case 2:    sExecCmd = "Dir """ & sTrgtDir & """ /b /s /a:d > """ & sTmpFilePath & """"
        Case Else: sExecCmd = ""
    End Select
    With CreateObject("Wscript.Shell")
        .Run "cmd /c" & sExecCmd, 7, True
    End With
    
    Dim objFile As Object
    Dim sTextAll As String
    On Error Resume Next
    If Err.Number = 0 Then
        Set objFile = objFSO.OpenTextFile(sTmpFilePath, 1)
        If Err.Number = 0 Then
            sTextAll = objFile.ReadAll
            sTextAll = Left(sTextAll, Len(sTextAll) - Len(vbNewLine))       '�����ɉ��s���t�^����Ă��܂����߁A�폜
            asFileList = Split(sTextAll, vbNewLine)
            objFile.Close
        Else
            MsgBox "�t�@�C�����J���܂���: " & Err.Description
        End If
        Set objFile = Nothing   '�I�u�W�F�N�g�̔j��
    Else
        MsgBox "�G���[ " & Err.Description
    End If
    objFSO.DeleteFile sTmpFilePath, True
    Set objFSO = Nothing    '�I�u�W�F�N�g�̔j��
    On Error GoTo 0
End Function
    Private Sub Test_GetFileList2()
        Dim objWshShell As Object
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sCurDir As String
        sCurDir = "C:\codes"
        
        Dim asFileList As Variant
'        Call GetFileList2("Z:\300_Musics", asFileList, 0)
'        Call GetFileList2("Z:\300_Musics", asFileList, 1)
        Call GetFileList2("Z:\300_Musics", asFileList, 2)
    End Sub

Public Function GetFileOrFolder( _
    ByVal sChkTrgtPath As String _
) As T_SYSOBJ_TYPE
    Dim oFileSys As Object
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    If oFileSys.FolderExists(sChkTrgtPath) = True And _
       oFileSys.FileExists(sChkTrgtPath) = False Then
        GetFileOrFolder = SYSOBJ_DIRECTORY
    Else
        If oFileSys.FolderExists(sChkTrgtPath) = False And _
           oFileSys.FileExists(sChkTrgtPath) = True Then
            GetFileOrFolder = SYSOBJ_FILE
        Else
            GetFileOrFolder = SYSOBJ_NOT_EXIST
        End If
    End If
    Set oFileSys = Nothing
End Function

' ==================================================================
' = �T�v    �t�H���_�I���_�C�A���O��\������
' = ����    sInitPath   String  [in]  �f�t�H���g�t�H���_�p�X�i�ȗ��j
' = �ߒl                String        �t�H���_�I������
' = �o��    �Ȃ�
' ==================================================================
Private Function ShowFolderSelectDialog( _
    Optional ByVal sInitPath As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFolderPicker)
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
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then '�L�����Z������
        ShowFolderSelectDialog = ""
    Else
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FolderExists(sSelectedPath) Then
            ShowFolderSelectDialog = sSelectedPath
        Else
            ShowFolderSelectDialog = ""
        End If
    End If
    
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFolderSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        MsgBox ShowFolderSelectDialog( _
                    objWshShell.SpecialFolders("Desktop") _
                )
    End Sub

' ==================================================================
' = �T�v    �t�@�C���i�P��j�I���_�C�A���O��\������
' = ����    sInitPath   String  [in]  �f�t�H���g�t�@�C���p�X�i�ȗ��j
' = ����    sFilters�@  String  [in]  �I�����̃t�B���^�i�ȗ��j(��)
' = �ߒl                String        �t�@�C���I������
' = �o��    (��)�_�C�A���O�̃t�B���^�w����@�͈ȉ��B
' =              ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =                    �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =                    �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =                    �E�t�B���^����������ꍇ�A","�ŋ�؂�
' =         sFilters ���ȗ��������͋󕶎��̏ꍇ�A�t�B���^���N���A����B
' ==================================================================
Private Function ShowFileSelectDialog( _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFilters As String = "" _
) As String
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "�t�@�C����I�����Ă�������"
    fdDialog.AllowMultiSelect = False
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) '�t�B���^�ǉ�
 
    '�_�C�A���O�\��
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then '�L�����Z������
        ShowFileSelectDialog = ""
    Else
        Dim sSelectedPath As String
        sSelectedPath = fdDialog.SelectedItems.Item(1)
        If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
            ShowFileSelectDialog = sSelectedPath
        Else
            ShowFileSelectDialog = ""
        End If
    End If
 
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFileSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sFilters As String
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png"
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv"
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png,�e�L�X�g�t�@�C��/*.txt; *.csv"
        sFilters = ""
        
        MsgBox ShowFileSelectDialog( _
                    objWshShell.SpecialFolders("Desktop") & "\test.txt", _
                    sFilters _
                )
    '    MsgBox ShowFileSelectDialog( _
    '                objWshShell.SpecialFolders("Desktop") & "\test.txt" _
    '            )
    End Sub
 
' ==================================================================
' = �T�v    �t�@�C���i�����j�I���_�C�A���O��\������
' = ����    asSelectedFiles String()    [out] �I�����ꂽ�t�@�C���p�X�ꗗ
' = ����    sInitPath       String      [in]  �f�t�H���g�t�@�C���p�X�i�ȗ��j
' = ����    sFilters�@      String      [in]  �I�����̃t�B���^�i�ȗ��j(��)
' = �ߒl    �Ȃ�
' = �o��    (��)�_�C�A���O�̃t�B���^�w����@�͈ȉ��B
' =              ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
' =                    �E�g���q����������ꍇ�́A";"�ŋ�؂�
' =                    �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
' =                    �E�t�B���^����������ꍇ�A","�ŋ�؂�
' =         sFilters ���ȗ��������͋󕶎��̏ꍇ�A�t�B���^���N���A����B
' ==================================================================
Private Function ShowFilesSelectDialog( _
    ByRef asSelectedFiles() As String, _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFilters As String = "" _
)
    Dim fdDialog As Office.FileDialog
    Set fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.Title = "�t�@�C����I�����Ă��������i�����j"
    fdDialog.AllowMultiSelect = True
    If sInitPath = "" Then
        'Do Nothing
    Else
        fdDialog.InitialFileName = sInitPath
    End If
    Call SetDialogFilters(sFilters, fdDialog) '�t�B���^�ǉ�
 
    '�_�C�A���O�\��
    Dim lResult As Long
    lResult = fdDialog.Show()
    If lResult <> -1 Then '�L�����Z������
        ReDim Preserve asSelectedFiles(0)
        asSelectedFiles(0) = ""
    Else
        Dim lSelNum As Long
        lSelNum = fdDialog.SelectedItems.Count
        ReDim Preserve asSelectedFiles(lSelNum - 1)
        Dim lSelIdx As Long
        For lSelIdx = 0 To lSelNum - 1
            Dim sSelectedPath As String
            sSelectedPath = fdDialog.SelectedItems(lSelIdx + 1)
            If CreateObject("Scripting.FileSystemObject").FileExists(sSelectedPath) Then
                asSelectedFiles(lSelIdx) = sSelectedPath
            Else
                asSelectedFiles(lSelIdx) = ""
            End If
        Next lSelIdx
    End If
 
    Set fdDialog = Nothing
End Function
    Private Sub Test_ShowFilesSelectDialog()
        Dim objWshShell
        Set objWshShell = CreateObject("WScript.Shell")
        Dim sFilters As String
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png"
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv"
        'sFilters = "�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png,�e�L�X�g�t�@�C��/*.txt; *.csv"
        sFilters = "�S�Ẵt�@�C��/*.*,�摜�t�@�C��/*.gif; *.jpg; *.jpeg; *.png,�e�L�X�g�t�@�C��/*.txt; *.csv"
 
        Dim asSelectedFiles() As String
        Call ShowFilesSelectDialog( _
                    asSelectedFiles, _
                    objWshShell.SpecialFolders("Desktop") & "\test.txt", _
                    sFilters _
                )
        Dim sBuf As String
        sBuf = ""
        sBuf = sBuf & vbNewLine & UBound(asSelectedFiles) + 1
        Dim lSelIdx As Long
        For lSelIdx = 0 To UBound(asSelectedFiles)
            sBuf = sBuf & vbNewLine & asSelectedFiles(lSelIdx)
        Next lSelIdx
        MsgBox sBuf
    End Sub
 
'ShowFileSelectDialog() �� ShowFilesSelectDialog() �p�̊֐�
'�_�C�A���O�̃t�B���^��ǉ�����B�w����@�͈ȉ��B
'  ex) �摜�t�@�C��/*.gif; *.jpg; *.jpeg,�e�L�X�g�t�@�C��/*.txt; *.csv
'      �E�g���q����������ꍇ�́A";"�ŋ�؂�
'      �E�t�@�C����ʂƊg���q��"/"�ŋ�؂�
'      �E�t�B���^����������ꍇ�A","�ŋ�؂�
'sFilters ���󕶎��̏ꍇ�A�t�B���^���N���A����B
Private Function SetDialogFilters( _
    ByVal sFilters As String, _
    ByRef fdDialog As FileDialog _
)
    fdDialog.Filters.Clear
    If sFilters = "" Then
        'Do Nothing
    Else
        Dim vFilter As Variant
        If InStr(sFilters, ",") > 0 Then
            Dim vFilters As Variant
            vFilters = Split(sFilters, ",")
            Dim lFilterIdx As Long
            For lFilterIdx = 0 To UBound(vFilters)
                If InStr(vFilters(lFilterIdx), "/") > 0 Then
                    vFilter = Split(vFilters(lFilterIdx), "/")
                    If UBound(vFilter) = 1 Then
                        fdDialog.Filters.Add vFilter(0), vFilter(1), lFilterIdx + 1
                    Else
                        MsgBox _
                            "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                            """/"" �͈�����w�肵�Ă�������" & vbNewLine & _
                            "  " & vFilters(lFilterIdx)
                        MsgBox "�����𒆒f���܂��B"
                        End
                    End If
                Else
                    MsgBox _
                        "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                        "��ʂƊg���q�� ""/"" �ŋ�؂��Ă��������B" & vbNewLine & _
                        "  " & vFilters(lFilterIdx)
                    MsgBox "�����𒆒f���܂��B"
                    End
                End If
            Next lFilterIdx
        Else
            If InStr(sFilters, "/") > 0 Then
                vFilter = Split(sFilters, "/")
                If UBound(vFilter) = 1 Then
                    fdDialog.Filters.Add vFilter(0), vFilter(1), 1
                Else
                    MsgBox _
                        "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                        """/"" �͈�����w�肵�Ă�������" & vbNewLine & _
                        "  " & sFilters
                    MsgBox "�����𒆒f���܂��B"
                    End
                End If
            Else
                MsgBox _
                    "�t�@�C���I���_�C�A���O�̃t�B���^�̎w����@������Ă��܂�" & vbNewLine & _
                    "��ʂƊg���q�� ""/"" �ŋ�؂��Ă��������B" & vbNewLine & _
                    "  " & sFilters
                MsgBox "�����𒆒f���܂��B"
                End
            End If
        End If
    End If
End Function
