'TODO:�v�R�}���h��

Option Explicit
'==========================================================
'= �C���N���[�h
'==========================================================
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetDetailsOfGetDetailsOf()

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
