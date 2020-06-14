' ******************************************************************
' * �T�v    �A�h�C���p�̐ݒ�t�@�C�����Ǘ�����N���X
' *
' * �o��    �ݒ�͕�����^�Ƃ��ĕۑ�����B���̂��߈ȉ��ɒ��ӂ��邱��
' *           - Boolean�l�͕�������w�肷�邱��(ex.True�ł͂Ȃ�"True")
' *           - ���䕶���͐��䕶����������������w�肷�邱��(ex.vbTab�ł͂Ȃ�"vbTab")
' ******************************************************************

' ******************************************************************
' * �{����
' ******************************************************************
Class SettingFile
    Private gdSettingItems
    Private gsSettingFilePath
    Private gsDelimiter

    ' ==================================================================
    ' = �T�v    �R���X�g���N�^
    ' ==================================================================
    Private Sub Class_Initialize()
        Call Init
    End Sub

    ' ==================================================================
    ' = �T�v    �f�X�g���N�^
    ' ==================================================================
    Private Sub Class_Terminate()
        Call DeInit
    End Sub

    ' ==================================================================
    ' = �T�v    ����������
    ' ==================================================================
    Private Sub Init()
        Set gdSettingItems = CreateObject("Scripting.Dictionary")
        gsSettingFilePath = ""
        gsDelimiter = vbTab
    End Sub

    ' ==================================================================
    ' = �T�v    �I������
    ' ==================================================================
    Private Sub DeInit()
        Set gdSettingItems = Nothing
        gsSettingFilePath = ""
        gsDelimiter = vbTab
    End Sub

    ' ==================================================================
    ' = �T�v    �t�@�C����ǂݏo��
    ' = ����    sFilePath       String      [in]    �t�@�C���p�X
    ' = ����    sDelimiter      String      [in]    �f���~�^
    ' = �ߒl                    Boolean             �ǂݏo������
    ' = �o��    �E�u�ǂݏo���t�@�C���̋�؂蕶���v��sDelimiter����v�����邱��
    ' =           ��v���Ȃ��ꍇ�A�����𒆒f����B
    ' =         �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - sFilePath�����݂��Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function FileLoad( _
        ByVal sFilePath, _
        ByVal sDelimiter _
    )
        gsDelimiter = sDelimiter
        
        gsSettingFilePath = sFilePath
        'Debug.Print gsSettingFilePath
        
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        If objFSO.FileExists(gsSettingFilePath) Then
            Dim vFileLineAll()
            
            Dim objTxtFile
            Set objTxtFile = objFSO.OpenTextFile(gsSettingFilePath, 1, True)
            FileLoad = False
            Do Until objTxtFile.AtEndOfStream
                Dim vKeyValue
                Dim sLine
                sLine = objTxtFile.ReadLine
                If InStr(sLine, gsDelimiter) Then
                    vKeyValue = Split(sLine, gsDelimiter)
                    If UBound(vKeyValue) = 0 Then
                        gdSettingItems.Add vKeyValue(0), ""           '�P���؂蕶��(�l�Ȃ�)
                    ElseIf UBound(vKeyValue) = 1 Then
                        gdSettingItems.Add vKeyValue(0), vKeyValue(1) '�P���؂蕶��(�l����)
                    Else
                        Stop                                          '������؂蕶��
                    End If
                Else
                    Stop                                              '��؂蕶���Ȃ�
                End If
                FileLoad = True
            Loop
            objTxtFile.Close
        Else
            FileLoad = False
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �t�@�C���������o��
    ' = ����    �Ȃ�
    ' = �ߒl                    Boolean             �����o������
    ' = �o��    �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - ���O��FileLoad()���Ă΂�Ă��Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function FileSave()
        If gdSettingItems Is Nothing Then
            FileSave = False
        Else
            'Debug.Print gsSettingFilePath
            
            Dim objFSO
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Dim objTxtFile
            Set objTxtFile = objFSO.OpenTextFile(gsSettingFilePath, 2, True)
            Dim vKey
            For Each vKey In gdSettingItems
                objTxtFile.WriteLine vKey & gsDelimiter & gdSettingItems.Item(vKey)
            Next
            objTxtFile.Close
            FileSave = True
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ��ǉ�����
    ' = ����    sKey            String      [in]    �ݒ�L�[
    ' = ����    sValue          String      [in]    �ݒ�l
    ' = ����    bDoSave         String      [in]    �t�@�C���ۑ����{�L��
    ' = �ߒl                    Boolean             �ǉ�����
    ' = �o��    �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - ���O��FileLoad()���Ă΂�Ă��Ȃ�
    ' =           - �ۑ��Ɏ��s
    ' = �ˑ�    Me/FileSave()
    ' ==================================================================
    Public Function Add( _
        ByVal sKey, _
        ByVal sValue, _
        ByVal bDoSave _
    )
        If gdSettingItems Is Nothing Then
            Add = False
        Else
            '�ǉ�
            If gdSettingItems.Exists(sKey) Then
                gdSettingItems.Item(sKey) = sValue
            Else
                gdSettingItems.Add sKey, sValue
            End If
            
            '�t�@�C���ۑ�
            If bDoSave = True Then
                Add = FileSave()
            Else
                Add = True
            End If
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ���폜����
    ' = ����    sKey            String      [in]    �ݒ�L�[
    ' = ����    bDoSave         String      [in]    �t�@�C���ۑ����{�L��
    ' = �ߒl                    Boolean             �폜����
    ' = �o��    �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - sKey�����݂��Ȃ�
    ' =           - ���O��FileLoad()���Ă΂�Ă��Ȃ�
    ' =           - �ۑ��Ɏ��s
    ' = �ˑ�    Me/FileSave()
    ' ==================================================================
    Public Function Delete( _
        ByVal sKey, _
        ByVal bDoSave _
    )
        If gdSettingItems Is Nothing Then
            Delete = False
        Else
            If gdSettingItems.Exists(sKey) Then
                '�폜
                gdSettingItems.Remove (sKey)
                
                '�t�@�C���ۑ�
                If bDoSave = True Then
                    Delete = FileSave()
                Else
                    Delete = True
                End If
            Else
                Delete = False
            End If
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ��S�폜����
    ' = ����    bDoSave         String      [in]    �t�@�C���ۑ����{�L��(�ȗ���)
    ' = �ߒl                    Boolean             �폜����
    ' = �o��    �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - ���O��FileLoad()���Ă΂�Ă��Ȃ�
    ' =           - �ۑ��Ɏ��s
    ' = �ˑ�    Me/FileSave()
    ' ==================================================================
    Public Function DeleteAll( _
        ByVal bDoSave _
    )
        If gdSettingItems Is Nothing Then
            DeleteAll = False
        Else
            '�S�폜
            gdSettingItems.RemoveAll
            
            '�t�@�C���ۑ�
            If bDoSave = True Then
                DeleteAll = FileSave()
            Else
                DeleteAll = True
            End If
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ�̑��݊m�F���s��
    ' = ����    sKey            String      [in]    �ݒ�L�[
    ' = �ߒl                    Boolean             ���݊m�F����
    ' = �o��    �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - ���O��FileLoad()���Ă΂�Ă��Ȃ�
    ' =           - sKey�����݂��Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function Exists( _
        ByVal sKey _
    )
        If gdSettingItems Is Nothing Then
            Exists = False
        Else
            If gdSettingItems.Exists(sKey) Then
                Exists = True
            Else
                Exists = False
            End If
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ萔���擾����
    ' = ����    �Ȃ�
    ' = �ߒl                    Long    �ݒ萔
    ' = �o��    �E�ȉ��̏ꍇ�A0��ԋp����
    ' =           - ���O��FileLoad()���Ă΂�Ă��Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Property Get Count()
        If gdSettingItems Is Nothing Then
            Count = 0
        Else
            Count = gdSettingItems.Count
        End If
    End Property

    ' ==================================================================
    ' = �T�v    �ݒ�l���擾����
    ' = ����    sKey            String      [in]    �ݒ�L�[
    ' = ����    sValue          String      [out]   �ݒ�l
    ' = �ߒl                    Boolean             �擾����
    ' = �o��    �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - ���O��FileLoad()���Ă΂�Ă��Ȃ�
    ' =           - sKey�����݂��Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function Item( _
        ByVal sKey, _
        ByRef sValue _
    )
        If gdSettingItems.Exists(sKey) Then
            sValue = gdSettingItems.Item(sKey)
            Item = True
        Else
            sValue = ""
            Item = False
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ�L�[�Ɛݒ�l��S�Ď擾����
    ' = ����    �Ȃ�
    ' = �ߒl                    Object(Dictionary)  �ݒ�L�[/�l����
    ' = �o��    �Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Property Get AllItems()
        Set AllItems = gdSettingItems
    End Property

    ' ==================================================================
    ' = �T�v    �t�@�C���I�[�v������ݒ�擾�i���Ȃ���ΐݒ�ǉ��j���ꊇ�ōs��
    ' = ����    sFilePath       String      [in]    �t�@�C���p�X
    ' = ����    sKey            String      [in]    �ݒ�L�[
    ' = ����    sValue          String      [out]   �ݒ�l
    ' = ����    sInitValue      String      [in]    �ݒ菉���l
    ' = ����    bDoSave         String      [in]    �t�@�C���ۑ����{�L��
    ' = �ߒl                    Boolean             �擾����
    ' = �o��    �E�t�@�C���I�[�v����A�ݒ�l���擾����B
    ' =           �ݒ�l�����݂��Ȃ��ꍇ�A�����l�Ƃ��Đݒ�l���X�V����B
    ' =         �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - sFilePath�����݂��Ȃ�
    ' =           - sKey�����݂��Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function ReadItemFromFile( _
        ByVal sFilePath, _
        ByVal sKey, _
        ByRef sValue, _
        ByVal sInitValue, _
        ByVal bDoSave _
    )
        Call Init
        
        '�ݒ�t�@�C���ǂݏo��
        Dim bExistFile
        Dim bExistItem
        bExistFile = Me.FileLoad(sFilePath, vbTab )
        
        '�ݒ荀�ڎ擾���X�V
        Dim sItem
        If bExistFile = True Then
            bExistItem = Me.Item(sKey, sItem)
            If bExistItem = True Then
                sValue = sItem
                'Call Me.Add(sKey, sValue, bDoSave)
                ReadItemFromFile = True
            Else
                sValue = sInitValue
                Call Me.Add(sKey, sValue, bDoSave)
                ReadItemFromFile = False
            End If
        Else
            sValue = sInitValue
            Call Me.Add(sKey, sValue, bDoSave)
            ReadItemFromFile = False
        End If
        
        Call DeInit
    End Function

    ' ==================================================================
    ' = �T�v    �t�@�C���I�[�v������ݒ�X�V/�ǉ����ꊇ�ōs��
    ' = ����    sFilePath       String      [in]    �t�@�C���p�X
    ' = ����    sKey            String      [in]    �ݒ�L�[
    ' = ����    sValue          String      [in]    �ݒ�l
    ' = �ߒl                    Boolean             �擾����
    ' = �o��    �E�t�@�C���I�[�v����A�ݒ�l���X�V/�ǉ�����B
    ' =         �E�ȉ��̏ꍇ�AFalse��ԋp����
    ' =           - sFilePath�����݂��Ȃ�
    ' =           - sKey�����݂��Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function WriteItemToFile( _
        ByVal sFilePath, _
        ByVal sKey, _
        ByVal sValue _
    )
        Call Init
        
        '�ݒ�t�@�C���ǂݏo��
        Dim bExistFile
        Dim bExistItem
        bExistFile = Me.FileLoad(sFilePath, vbTab)
        bExistItem = Me.Add(sKey, sValue, True)
        If bExistFile = True And bExistItem = True Then
            WriteItemToFile = True
        Else
            WriteItemToFile = False
        End If
        
        Call DeInit
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ�l�ϊ��p ������to�^�U�l�ϊ�
    ' = ����    sValue          String      [in]    �l(������)
    ' = �ߒl                    Boolean             �l(�^�U�l)
    ' = �o��    �Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function ConvTypeStr2Bool( _
        ByVal sValue _
    )
        If sValue = "True" Then
            ConvTypeStr2Bool = True
        Else
            ConvTypeStr2Bool = False
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ�l�ϊ��p �^�U�lto������ϊ�
    ' = ����    bValue          Boolean     [in]    �l(�^�U�l)
    ' = �ߒl                    String              �l(������)
    ' = �o��    �Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function ConvTypeBool2Str( _
        ByVal bValue _
    )
        If bValue = True Then
            ConvTypeBool2Str = "True"
        Else
            ConvTypeBool2Str = "False"
        End If
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ�l�ϊ��p ���lto���䕶�� �ϊ�
    ' = ����    sValue          String      [in]    �l(������)
    ' = �ߒl                    String              �l(���䕶��)
    ' = �o��    �Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function ConvStrRaw2CntrlChr( _
        ByVal sValue _
    )
        Select Case sValue
            Case "vbTab":     ConvStrRaw2CntrlChr = vbTab
            Case "vbCr":      ConvStrRaw2CntrlChr = vbCr
            Case "vbLf":      ConvStrRaw2CntrlChr = vbLf
            Case "vbNewLine": ConvStrRaw2CntrlChr = vbNewLine
            Case Else:        ConvStrRaw2CntrlChr = sValue
        End Select
    End Function

    ' ==================================================================
    ' = �T�v    �ݒ�l�ϊ��p ���䕶��to���l �ϊ�
    ' = ����    sValue          String      [in]    �l(���䕶��)
    ' = �ߒl                    String              �l(������)
    ' = �o��    �Ȃ�
    ' = �ˑ�    �Ȃ�
    ' ==================================================================
    Public Function ConvStrCntrlChr2Raw( _
        ByVal sValue _
    )
        Select Case sValue
            Case vbTab:     ConvStrCntrlChr2Raw = "vbTab"
            Case vbCr:      ConvStrCntrlChr2Raw = "vbCr"
            Case vbLf:      ConvStrCntrlChr2Raw = "vbLf"
            Case vbNewLine: ConvStrCntrlChr2Raw = "vbNewLine"
            Case Else:      ConvStrCntrlChr2Raw = sValue
        End Select
    End Function
End Class

    'Call Test_SettingFile()
    Private Sub Test_SettingFile()
        Const sTEST_KEY = "init value"
        Const sTEMP_FILE_NAME = "SettingFileClass.cfg"
        
        Dim clSetting
        Set clSetting = New SettingFile
        
        Dim sSettingFilePath
        sSettingFilePath = "C:\Users\" & CreateObject("WScript.Network").UserName & "\AppData\Local\Temp\" & sTEMP_FILE_NAME
        Dim sTestValue
        Call clSetting.ReadItemFromFile(sSettingFilePath, "sTEST_KEY", sTestValue, sTEST_KEY, False)
        sTestValue = InputBox("msg", "title", sTestValue)
        Call clSetting.WriteItemToFile(sSettingFilePath, "sTEST_KEY", sTestValue)
    End Sub
