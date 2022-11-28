Option Explicit

'<<�T�v>>
'  �t�@�C�������擾����B
'
'<<�g�p���@>>
'  GetFileDetailInfo.vbs <target_file_path> <info_type>...
'   �E�t�@�C���p�X�i<target_file_path>�j
'   �E�t�@�C������ʁi<info_type>�j(��1)
'     (��1) �t�@�C�������
'         [����]  [����]                  [�v���p�e�B��]      [�f�[�^�^]              [Get/Set]   [�o�͗�]
'         1       �t�@�C����              Name                vbString    ������^    Get/Set     03 Ride Featuring Tony Matterhorn.MP3
'         2       �t�@�C���T�C�Y          Size                vbLong      �������^    Get         4286923
'         3       �t�@�C�����            Type                vbString    ������^    Get         MPEG layer 3
'         4       �t�@�C���i�[��h���C�u  Drive               vbString    ������^    Get         Z:
'         5       �t�@�C���p�X            Path                vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
'         6       �e�t�H���_              ParentFolder        vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
'         7       MS-DOS�`���t�@�C����    ShortName           vbString    ������^    Get         03 Ride Featuring Tony Matterhorn.MP3
'         8       MS-DOS�`���p�X          ShortPath           vbString    ������^    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
'         9       �쐬����                DateCreated         vbDate      ���t�^      Get         2015/08/19 0:54:45
'         10      �A�N�Z�X����            DateLastAccessed    vbDate      ���t�^      Get         2016/10/14 6:00:30
'         11      �X�V����                DateLastModified    vbDate      ���t�^      Get         2016/10/14 6:00:30
'         12      ����                    Attributes          vbLong      �������^    (��2)       32
'     (��2) ����
'         [�l]                [����]                                      [������]    [Get/Set]
'         1  �i0b00000001�j   �ǂݎ���p�t�@�C��                        ReadOnly    Get/Set
'         2  �i0b00000010�j   �B���t�@�C��                                Hidden      Get/Set
'         4  �i0b00000100�j   �V�X�e���E�t�@�C��                          System      Get/Set
'         8  �i0b00001000�j   �f�B�X�N�h���C�u�E�{�����[���E���x��        Volume      Get
'         16 �i0b00010000�j   �t�H���_�^�f�B���N�g��                      Directory   Get
'         32 �i0b00100000�j   �O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1   Archive     Get/Set
'         64 �i0b01000000�j   �����N�^�V���[�g�J�b�g                      Alias       Get
'         128�i0b10000000�j   ���k�t�@�C��                                Compressed  Get

'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )  'GetFileInfo()

'===============================================================================
'= �ݒ�l
'===============================================================================
Const bEXEC_TEST = False '�e�X�g�p
Const sSCRIPT_NAME = "�t�@�C�����擾"

'===============================================================================
'= �{����
'===============================================================================
Dim cArgs '{{{
Set cArgs = CreateObject("System.Collections.ArrayList")

If bEXEC_TEST = True Then
    Call Test_Main()
Else
    Dim vArg
    For Each vArg in WScript.Arguments
        cArgs.Add vArg
    Next
    Call Main()
End If '}}}

'===============================================================================
'= ���C���֐�
'===============================================================================
Public Sub Main()
    If cArgs.Count >= 2 then
        Dim sTrgtPath
        sTrgtPath = cArgs(0)
        Dim cInfoTypes
        Set cInfoTypes = CreateObject("System.Collections.ArrayList")
        Dim lArgIdx
        For lArgIdx = 1 To (cArgs.Count - 1)
            If IsNumeric(cArgs(lArgIdx)) Then
                cInfoTypes.Add cArgs(lArgIdx)
            Else
                WScript.Echo "[error] <info_type> must be a number."
                WScript.Echo "  usage : GetFileDetailInfo.vbs <target_file_path> <info_type>..."
                Exit Sub
            End If
        Next
    Else
        WScript.Echo "[error] wrong number of argments"
        WScript.Echo "  usage : GetFileDetailInfo.vbs <target_file_path> <info_type>..."
        Exit Sub
    End If
    
    Dim vFileInfo
    Dim sOutputStr
    sOutputStr = sTrgtPath
    Dim sInfoType
    For Each sInfoType In cInfoTypes
        Dim bResult
        bResult = GetFileInfo( sTrgtPath, CLng(sInfoType), vFileInfo)
        If bResult = True Then
            sOutputStr = sOutputStr & vbTab & vFileInfo
        Else
            WScript.Echo "[error] GetFileInfo() failed."
            Exit Sub
        End If
    Next
    WScript.Echo sOutputStr
End Sub

'===============================================================================
'= �����֐�
'===============================================================================

'===============================================================================
'= �e�X�g�֐�
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTESTCASE_STRT = 1
    Const lTESTCASE_LAST = 6
    Dim lIdx
    For lIdx = lTESTCASE_STRT To lTESTCASE_LAST
        Dim sTestFuncName
        sTestFuncName = _
            "Test_Case" & _
            String(3 - Len(CStr(lIdx)), "0") & _
            CStr(lIdx)
        cArgs.Clear
        Dim oFuncPtr
        Set oFuncPtr = GetRef(sTestFuncName)
        WScript.Echo "=== " & sTestFuncName & " ==="
        oFuncPtr()
    Next
End Sub
Private Sub Test_Case001()
    
    cArgs.Add WScript.ScriptFullName
    cArgs.Add 1 '�t�@�C����
    cArgs.Add 2 '�t�@�C���T�C�Y
    Call Main()
    
End Sub
Private Sub Test_Case002()
    cArgs.Add WScript.ScriptFullName
    cArgs.Add 1
    cArgs.Add 1
    cArgs.Add 1
    cArgs.Add 1
    Call Main()
End Sub
Private Sub Test_Case003()
    cArgs.Add WScript.ScriptFullName
    Call Main()
End Sub
Private Sub Test_Case004()
    cArgs.Add WScript.ScriptFullName
    cArgs.Add "aaa"
    Call Main()
End Sub
Private Sub Test_Case005()
    cArgs.Add WScript.ScriptFullName
    cArgs.Add 13
    Call Main()
End Sub
Private Sub Test_Case006()
    cArgs.Add "aaa"
    cArgs.Add 1
    Call Main()
End Sub
'}}}

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile ) '{{{
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function '}}}

