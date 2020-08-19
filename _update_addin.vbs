Const sAddinDirPath = "C:\codes\vba\excel\AddIns"

'�t�H���_�p�X�ꗗ�擾
Dim cDirPathList
Set cDirPathList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sAddinDirPath, cDirPathList, 2, "")

''���f�o�b�O�p��
'Dim sDebugMsg
'sDebugMsg = ""
'Dim vDirPath
'For Each vDirPath In cDirPathList
'    sDebugMsg = sDebugMsg & vbNewLine & vDirPath
'Next
'MsgBox sDebugMsg
'WScript.Quit

'�R�s�[���t�H���_����
Dim sSrcDirPath
Dim vDirPathTmp
For Each vDirPathTmp In cDirPathList
    Dim oRegExp
    Dim sTargetStr
    Dim sSearchPattern
    Set oRegExp = CreateObject("VBScript.RegExp")
    sTargetStr = vDirPathTmp
    sSearchPattern = "\Tmp\d{8}$"
    oRegExp.Pattern = sSearchPattern
    oRegExp.IgnoreCase = True
    oRegExp.Global = True
    Dim oMatchResult
    Set oMatchResult = oRegExp.Execute(sTargetStr)
    If oMatchResult.Count > 0 Then
        sSrcDirPath = vDirPathTmp
        Exit For
    End If
Next

''���f�o�b�O�p��
'MsgBox sSrcDirPath
'WScript.Quit

'�R�s�[���t�H���_���̃t�@�C�����X�g�擾
Dim cSrcFilePathList
Set cSrcFilePathList = CreateObject("System.Collections.ArrayList")
Call GetFileListCmdClct(sSrcDirPath, cSrcFilePathList, 1, "")

''���f�o�b�O�p��
'Dim sDebugMsg
'sDebugMsg = ""
'For Each vSrcFilePath In cSrcFilePathList
'    sDebugMsg = sDebugMsg & vbNewLine & vSrcFilePath
'Next
'MsgBox sDebugMsg
'WScript.Quit

'�R�s�[���t�H���_���̃t�@�C�����R�s�[��t�H���_�փR�s�[
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Dim sDebugMsg '���f�o�b�O�p��
'sDebugMsg = "" '���f�o�b�O�p��
Dim vSrcFilePath
Dim sDstDirPath
sDstDirPath = sAddinDirPath & "\MyExcelAddin.bas\"
For Each vSrcFilePath In cSrcFilePathList
    objFSO.CopyFile vSrcFilePath, sDstDirPath
    'sDebugMsg = sDebugMsg & vbNewLine & vSrcFilePath & "��" & sDstDirPath '���f�o�b�O�p��
Next
'MsgBox sDebugMsg '���f�o�b�O�p��
'WScript.Quit '���f�o�b�O�p��

'�R�s�[���t�H���_�폜
objFSO.DeleteFolder sSrcDirPath, True

' ==================================================================
' = �T�v    �t�@�C��/�t�H���_�p�X�ꗗ���擾����(Collection,Dir�R�}���h��)
' = ����    sTrgtDir        String      [in]    �Ώۃt�H���_
' = ����    cFileList       Collections [out]   �t�@�C��/�t�H���_�p�X�ꗗ
' = ����    lFileListType   Long        [in]    �擾����ꗗ�̌`��
' =                                                 0�F����
' =                                                 1:�t�@�C��
' =                                                 2:�t�H���_
' =                                                 ����ȊO�F�i�[���Ȃ�
' = ����    sFileExtStr     String      [in]    �擾����t�@�C���̊g���q
' =                                                 ex1) ""
' =                                                 ex2) "*"
' =                                                 ex3) "*.c"
' =                                                 ex4) "*.txt *.log *.csv"
' = �ߒl    �Ȃ�
' = �o��    �EDir �R�}���h�ɂ��t�@�C���ꗗ�擾�BGetFileList() ���������B
' = �o��    �EsFileExtStr�̓t�@�C���w�莞�̂ݗL��
' = �ˑ�    �Ȃ�
' = ����    FileSystem.vbs
' ==================================================================
Public Function GetFileListCmdClct( _
    ByVal sTrgtDir, _
    ByRef cFileList, _
    ByVal lFileListType, _
    ByVal sFileExtStr _
)
    Dim objFSO  'FileSystemObject�̊i�[��
    Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    
    'Dir �R�}���h���s�i�o�͌��ʂ��ꎞ�t�@�C���Ɋi�[�j
    Dim sTmpFilePath
    Dim sExecCmd
    sTmpFilePath = WScript.CreateObject( "WScript.Shell" ).CurrentDirectory & "\Dir.tmp"
    Dim sTrgtDirStr
    If sFileExtStr = "" Then
        sTrgtDirStr = """" & sTrgtDir & """"
    Else
        Dim vFileExtentions
        vFileExtentions = Split(sFileExtStr, " ")
        Dim lSplitIdx
        For lSplitIdx = 0 To UBound(vFileExtentions)
            If sTrgtDirStr = "" Then
                sTrgtDirStr = """" & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            Else
                sTrgtDirStr = sTrgtDirStr & " """ & sTrgtDir & "\" & vFileExtentions(lSplitIdx) & """"
            End If
        Next
    End If
    Select Case lFileListType
        Case 0:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a > """ & sTmpFilePath & """"
        Case 1:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:a-d > """ & sTmpFilePath & """"
        Case 2:    sExecCmd = "Dir " & sTrgtDirStr & " /b /s /a:d > """ & sTmpFilePath & """"
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
'   Call Test_GetFileListCmdClct()
    Private Sub Test_GetFileListCmdClct()
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Dim sCurDir
        sCurDir = "C:\codes"
        
        Dim cFileList
        Set cFileList = CreateObject("System.Collections.ArrayList")
        Call GetFileListCmdClct( sCurDir, cFileList, 1, "*.c *.h" )
        'Call GetFileListCmdClct( sCurDir, cFileList, 1, "*.h" )
        'Call GetFileListCmdClct( sCurDir, cFileList, 1, "" )
        'Call GetFileListCmdClct( sCurDir, cFileList, 2, "" )
        
        dim sFilePath
        dim sOutput
        sOutput = ""
        for each sFilePath in cFileList
            sOutput = sOutput & vbNewLine & sFilePath
        next
        MsgBox sOutput
    End Sub
