# �t�@�C���E�t�H���_���
## �v���p�e�B�ꗗ
- �t�@�C�����iFile �I�u�W�F�N�g�j

| �v���p�e�B��     | ����                   |>         | �f�[�^�^        | Get/Set | �o�͗�                                                                                    |
|:---|:---|:---|:---|:---|:---|
| Name             | �t�@�C����             | vbString | ������^        | Get/Set | 03 Ride Featuring Tony Matterhorn.MP3                                                     |
| Size             | �t�@�C���T�C�Y         | vbLong   | �������^ (Long) | Get     | 4286923                                                                                   |
| Type             | �t�@�C�����           | vbString | ������^        | Get     | MPEG layer 3                                                                              |
| Drive            | �t�@�C���i�[��h���C�u | vbString | ������^        | Get     | Z:                                                                                        |
| Path             | �t�@�C���p�X           | vbString | ������^        | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3 |
| ParentFolder     | �e�t�H���_             | vbString | ������^        | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice                                       |
| ShortName        | MS-DOS�`���t�@�C����   | vbString | ������^        | Get     | 03 Ride Featuring Tony Matterhorn.MP3                                                     |
| ShortPath        | MS-DOS�`���p�X         | vbString | ������^        | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3 |
| DateCreated      | �쐬����               | vbDate   | ���t�^ (Date)   | Get     | 2015/08/19 0:54:45                                                                        |
| DateLastAccessed | �A�N�Z�X����           | vbDate   | ���t�^ (Date)   | Get     | 2016/10/14 6:00:30                                                                        |
| DateLastModified | �X�V����               | vbDate   | ���t�^ (Date)   | Get     | 2016/10/14 6:00:30                                                                        |
| Attributes       | ����                   | vbLong   | �������^ (Long) | (��)    | 32                                                                                        |

- �t�H���_���iFolder �I�u�W�F�N�g�j

| �v���p�e�B��     | ����                   |>          | �f�[�^�^           | Get/Set | �o�͗�                                              |
|:---|:---|:---|:---|:---|:---|
| Name             | �t�H���_��             | vbString  | ������^           | Get/Set | Sacrifice                                           |
| Size             | �t�H���_�T�C�Y         | vbLong    | �������^ (Long)    | Get     | 80613775                                            |
| Type             | �t�@�C�����           | vbString  | ������^           | Get     | �t�@�C�� �t�H���_�[                                 |
| Drive            | �t�@�C���i�[��h���C�u | vbString  | ������^           | Get     | Z:                                                  |
| Path             | �t�H���_�p�X           | vbString  | ������^           | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice |
| IsRootFolder     | ���[�g �t�H���_        | vbBoolean | �u�[���^ (Boolean) | Get     | False                                               |
| ShortName        | MS-DOS�`���t�@�C����   | vbString  | ������^           | Get     | Sacrifice                                           |
| ShortPath        | MS-DOS�`���p�X         | vbString  | ������^           | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice |
| DateCreated      | �쐬����               | vbDate    | ���t�^ (Date)      | Get     | 2015/08/19 0:54:44                                  |
| DateLastAccessed | �A�N�Z�X����           | vbDate    | ���t�^ (Date)      | Get     | 2015/08/19 0:54:44                                  |
| DateLastModified | �X�V����               | vbDate    | ���t�^ (Date)      | Get     | 2015/04/18 3:38:36                                  |
| Attributes       | ����                   | vbLong    | �������^ (Long)    | (��)    | 16                                                  |

- �����i���j

| ������     | ����                                      | Get/Set(��) | �r�b�g            |
|:---|:---|:---|:---|
| ReadOnly   | �ǂݎ���p�t�@�C��                      | Get/Set     | 1�i0b00000001�j   |
| Hidden     | �B���t�@�C��                              | Get/Set     | 2�i0b00000010�j   |
| System     | �V�X�e���E�t�@�C��                        | Get/Set     | 4�i0b00000100�j   |
| Volume     | �f�B�X�N�h���C�u�E�{�����[���E���x��      | Get         | 8�i0b00001000�j   |
| Directory  | �t�H���_�^�f�B���N�g��                    | Get         | 16�i0b00010000�j  |
| Archive    | �O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1 | Get/Set     | 32�i0b00100000�j  |
| Alias      | �����N�^�V���[�g�J�b�g                    | Get         | 64�i0b01000000�j  |
| Compressed | ���k�t�@�C��                              | Get         | 128�i0b10000000�j |

## ���s��
``` vba
Sub test()
	Dim sDirPath As String
	Dim sFileName As String
	Dim sFilePath As String
	sDirPath = "Z:\300_Musics\200_Reggae@Jamaica\Artist\Alaine\Sacrifice"
	sFileName = "03 Ride Featuring Tony Matterhorn.MP3"
	sFilePath = sDirPath & "\" & sFileName
	 
	Dim objFSO As Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	'=====================================================
	' �t�@�C�����
	'=====================================================
	Dim objFile As Object
	Set objFile = objFSO.GetFile(sFilePath)
	Debug.Print "�������t�@�C����񁖁���"
	Debug.Print "�y�t�@�C�����z" & objFile.Name
	Debug.Print "�y�t�@�C���T�C�Y�z" & objFile.Size
	Debug.Print "�E"
	Debug.Print "�E"
	Debug.Print "�E"
	Debug.Print ""
	
	'=====================================================
	' �t�H���_���
	'=====================================================
	Dim objFolder As Object
	Set objFolder = objFSO.GetFolder(sDirPath)
	Debug.Print "�������t�H���_��񁖁���"
	Debug.Print "�y�t�H���_���z" & objFolder.Name
	Debug.Print "�y�t�H���_�T�C�Y�z" & objFolder.Size
	Debug.Print "�E"
	Debug.Print "�E"
	Debug.Print "�E"
	Debug.Print ""
	
	'=====================================================
	' �g���b�N���
	'=====================================================
	Set objFolder = CreateObject("Shell.Application").Namespace(sDirPath & "\")
	
	'����t�@�C����ΏۂƂ���ꍇ
	Set objFile = objFolder.ParseName(sFileName)    '�t�@�C�������o��
	Debug.Print "�y�t�@�C���T�C�Y�z" & objFolder.GetDetailsOf(objFile, 1)   '�� �t�@�C���T�C�Y�F4.08 MB�i�t�@�C���T�C�Y�j
	Debug.Print "�E"
	Debug.Print "�E"
	Debug.Print "�E"
	Debug.Print ""
	
	'�t�H���_�����ׂẴt�@�C����ΏۂƂ���ꍇ
	For Each objFile In objFolder.Items
		Debug.Print "�y�t�@�C���T�C�Y�z" & objFolder.GetDetailsOf(objFile, 1)   '�� �t�@�C���T�C�Y�F4.08 MB�i�t�@�C���T�C�Y�j
		Debug.Print "�E"
		Debug.Print "�E"
		Debug.Print "�E"
	Next
	
	Set objFSO = Nothing
	Set objFolder = Nothing
	Set objFile = Nothing
End Sub
```

# �g���b�N���iGetDetailsOf�j
## �v���p�e�B�i�v���p�e�B�͂n�r�̃o�[�W�����ɂ���ĈقȂ�B�ȉ��̃R�[�h�Ŏ擾����B�j

| ������ | ����                   |>         | �f�[�^�^ | �o�͗�                          |
|:---|:---|:---|:---|:---|
| 1        | �t�@�C���T�C�Y         | vbString | ������^ | 4.08 MB                         |
| 2        | �t�@�C���̎��         | vbString | ������^ | MPEG layer 3                    |
| 3        | �X�V����               | vbString | ������^ | 2016/10/14 6:00                 |
| �E       | �E                     | �E       | �E       | �E                              |
| �E       | �E                     | �E       | �E       | �E                              |
| �E       | �E                     | �E       | �E       | �E                              |

## �v���p�e�B���擾�R�[�h
```vba
'GetDetailsOf()�̏ڍ׏��i�v�f�ԍ��A�^�C�g�����A�^���A�f�[�^�j���擾����
Public Sub GetDetailsOfGetDetailsOf()
	Dim sTrgtFolderPath As String
	Dim sTrgtFileName As String
	Dim sLogFilePath As String
	sTrgtFolderPath = "Z:\300_Musics\200_Reggae@Jamaica\Artist\Alaine\Sacrifice"
	sTrgtFileName = "03 Ride Featuring Tony Matterhorn.MP3"
	sLogFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\track_title_names.txt"
	
	Dim objFolder As Object
	Set objFolder = CreateObject("Shell.Application").Namespace(sTrgtFolderPath & "\")
	Dim objFile As Object
	Set objFile = objFolder.ParseName(sTrgtFileName)
	
	Open sLogFilePath For Output As #1
	Print #1, "[Idx] " & Chr(9) & "[TypeName]" & Chr(9) & "[Title]" & Chr(9) & "[Data]"
	Dim i As Long
	For i = 0 To 400
		Print #1, _
			i & Chr(9) & _
			TypeName(objFolder.GetDetailsOf(objFile, i)) & Chr(9) & _
			objFolder.GetDetailsOf("", i) & Chr(9) & _
			objFolder.GetDetailsOf(objFile, i)
	Next i
	Close #1
End Sub
```
