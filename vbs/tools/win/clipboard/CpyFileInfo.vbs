'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

' [ObjectPath]  [ObjectDirPath] [ObjectName]    [DateLastModified]  [DateCreated]   [DateLastAccessed]  [Size]  [Type]  [Attributes]

'####################################################################
'### �ݒ�
'####################################################################
Const INCLUDE_DOUBLE_QUOTATION = False

'####################################################################
'### ���O����
'####################################################################
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" ) 'GetFileOrFolder()
                                                   'GetFileInfo()
                                                   'GetFolderInfo()

'####################################################################
'### �{����
'####################################################################
Const sPROG_NAME = "�t�@�C���p�X/���O/�X�V�����R�s�["

Dim bIsContinue
bIsContinue = True

Dim cFilePaths

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        Dim sArg
        For Each sArg In WScript.Arguments
            cFilePaths.add sArg
        Next
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        Set cFilePaths = WScript.Col( WScript.Env("Selected") )
    Else '�f�o�b�O���s
        MsgBox "�f�o�b�O���[�h�ł��B"
        Set cFilePaths = CreateObject("System.Collections.ArrayList")
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\aabbbbb.txt"
        cFilePaths.Add "C:\Users\draem_000\Desktop\test\b b"
    End If
Else
    'Do Nothing
End If

'*** �t�@�C���p�X�`�F�b�N ***
If bIsContinue = True Then
    If cFilePaths.Count = 0 Then
        MsgBox "�t�@�C�����I������Ă��܂���", vbYes, sPROG_NAME
        MsgBox "�����𒆒f���܂�", vbYes, sPROG_NAME
        bIsContinue = False
    Else
        'Do Nothing
    End If
Else
    'Do Nothing
End If

'*** �N���b�v�{�[�h�փR�s�[ ***
If bIsContinue = True Then
    Dim sOutString
    Dim bFirstStore
    bFirstStore = True
    Dim oFilePath
    Dim oDirPath
    Dim sObjName
    Dim sModDate
    Dim sObjInfo
    Dim sObjSize
    Dim sObjType
    Dim sCreateDate
    Dim sAccessDate
    Dim sAttribute
    Dim lObjType
    For Each oFilePath In cFilePaths
        lObjType = GetFileOrFolder(oFilePath)
        Select case lObjType:
            Case 1 'File
                call GetFileInfo(oFilePath, 1, sObjName)
                call GetFileInfo(oFilePath, 2, sObjSize)
                call GetFileInfo(oFilePath, 3, sObjType)
                call GetFileInfo(oFilePath, 6, oDirPath)
                call GetFileInfo(oFilePath, 9, sCreateDate)
                call GetFileInfo(oFilePath, 10, sAccessDate)
                call GetFileInfo(oFilePath, 11, sModDate)
                call GetFileInfo(oFilePath, 12, sAttribute)
                sObjInfo = oFilePath & _
                    vbTab & oDirPath & _
                    vbTab & sObjName & _
                    vbTab & sModDate & _
                    vbTab & sCreateDate & _
                    vbTab & sAccessDate & _
                    vbTab & sObjSize & _
                    vbTab & sObjType & _
                    vbTab & sAttribute
            Case 2 'Folder
                call GetFolderInfo(oFilePath, 1, sObjName)
                call GetFolderInfo(oFilePath, 2, sObjSize)
                call GetFolderInfo(oFilePath, 3, sObjType)
                call GetFolderInfo(oFilePath, 9, sCreateDate)
                call GetFolderInfo(oFilePath, 10, sAccessDate)
                call GetFolderInfo(oFilePath, 11, sModDate)
                call GetFolderInfo(oFilePath, 12, sAttribute)
                oDirPath = Left(oFilePath, len(oFilePath) - len(sObjName) - 1)
                sObjInfo = oFilePath & _
                    vbTab & oDirPath & _
                    vbTab & sObjName & _
                    vbTab & sModDate & _
                    vbTab & sCreateDate & _
                    vbTab & sAccessDate & _
                    vbTab & sObjSize & _
                    vbTab & sObjType & _
                    vbTab & sAttribute
            Case Else 'Not Exist
                'Do Nohting
        End Select
        If bFirstStore = True Then
            sOutString = sObjInfo
            bFirstStore = False
        Else
            sOutString = sOutString & vbNewLine & sObjInfo
        End If
    Next
    CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
    'Do Nothing
End If

'####################################################################
'### �C���N���[�h�֐�
'####################################################################
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

