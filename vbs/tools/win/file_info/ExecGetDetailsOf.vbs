'TODO:�v�R�}���h��
'����́AGetDetailsOfGetDetailsOf()�̃e�X�g�p�X�N���v�g�̈ʒu�Â��B
'�t�@�C���p�X�������Ŏ󂯎��悤�ɂ���΁A�ėp���\�B
'���킹�ăR�}���h���������B

Option Explicit
'==========================================================
'= �C���N���[�h
'==========================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )  'GetDetailsOfGetDetailsOf()

'==========================================================
'= �{����
'==========================================================
'GetDetailsOf()�̏ڍ׏��i�v�f�ԍ��A�^�C�g�����A�^���A�f�[�^�j���擾����
Dim sTrgtFilePath
sTrgtFilePath = "Z:\300_Musics\200_Reggae@Jamaica\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3"

Dim sLogFilePath
sLogFilePath = WScript.CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\track_title_names.txt"

Call GetDetailsOfGetDetailsOf( sTrgtFilePath, sLogFilePath )

WScript.CreateObject("WScript.Shell").Run sLogFilePath, 1, True

'==========================================================
'= �C���N���[�h�֐�
'==========================================================
Private Function Include( ByVal sOpenFile )
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function

