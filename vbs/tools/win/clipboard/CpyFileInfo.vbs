'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

' [Path]	[Name]	[Size]	[Type]	[DateLastModified]	[DateCreated]	[DateLastAccessed]	[Attributes]

'####################################################################
'### �ݒ�
'####################################################################
Const INCLUDE_DOUBLE_QUOTATION = False

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "�t�@�C���p�X/���O/�X�V�����R�s�["

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
		MsgBox "�t�@�C�����I������Ă��܂���", vbYes, PROG_NAME
		MsgBox "�����𒆒f���܂�", vbYes, PROG_NAME
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
	Dim sFileName
	Dim sModDate
	Dim sFileInfo
	Dim sFileSize
	Dim sFileType
	Dim sCreateDate
	Dim sAccessDate
	Dim sAttribute
	For Each oFilePath In cFilePaths
		call GetFileInfo(oFilePath, 1, sFileName)
		call GetFileInfo(oFilePath, 2, sFileSize)
		call GetFileInfo(oFilePath, 3, sFileType)
		call GetFileInfo(oFilePath, 9, sCreateDate)
		call GetFileInfo(oFilePath, 10, sAccessDate)
		call GetFileInfo(oFilePath, 11, sModDate)
		call GetFileInfo(oFilePath, 12, sAttribute)
		sFileInfo = oFilePath & _
			vbTab & sFileName & _
			vbTab & sModDate & _
			vbTab & sCreateDate & _
			vbTab & sAccessDate & _
			vbTab & sFileSize & _
			vbTab & sFileType & _
			vbTab & sAttribute
		If bFirstStore = True Then
			sOutString = sFileInfo
			bFirstStore = False
		Else
			sOutString = sOutString & vbNewLine & sFileInfo
		End If
	Next
	CreateObject( "WScript.Shell" ).Exec( "clip" ).StdIn.Write( sOutString )
Else
	'Do Nothing
End If

' ==================================================================
' = �T�v	�t�@�C�����擾
' = ����	sTrgtPath		String		[in]	�t�@�C���p�X
' = ����	lGetInfoType	Long		[in]	�擾����� (��1)
' = ����	vFileInfo		Variant		[out]	�t�@�C����� (��1)
' = �ߒl					Boolean				�擾����
' = �o��	�ȉ��A�Q�ƁB
' =		(��1) �t�@�C�����
' =			[����]	[����]					[�v���p�e�B��]		[�f�[�^�^]				[Get/Set]	[�o�͗�]
' =			1		�t�@�C����				Name				vbString	������^	Get/Set		03 Ride Featuring Tony Matterhorn.MP3
' =			2		�t�@�C���T�C�Y			Size				vbLong		�������^	Get			4286923
' =			3		�t�@�C�����			Type				vbString	������^	Get			MPEG layer 3
' =			4		�t�@�C���i�[��h���C�u	Drive				vbString	������^	Get			Z:
' =			5		�t�@�C���p�X			Path				vbString	������^	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =			6		�e�t�H���_				ParentFolder		vbString	������^	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
' =			7		MS-DOS�`���t�@�C����	ShortName			vbString	������^	Get			03 Ride Featuring Tony Matterhorn.MP3
' =			8		MS-DOS�`���p�X			ShortPath			vbString	������^	Get			Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
' =			9		�쐬����				DateCreated			vbDate		���t�^		Get			2015/08/19 0:54:45
' =			10		�A�N�Z�X����			DateLastAccessed	vbDate		���t�^		Get			2016/10/14 6:00:30
' =			11		�X�V����				DateLastModified	vbDate		���t�^		Get			2016/10/14 6:00:30
' =			12		����					Attributes			vbLong		�������^	(��2)		32
' =		(��2) ����
' =			[�l]				[����]										[������]	[Get/Set]
' =			1  �i0b00000001�j	�ǂݎ���p�t�@�C��						ReadOnly	Get/Set
' =			2  �i0b00000010�j	�B���t�@�C��								Hidden		Get/Set
' =			4  �i0b00000100�j	�V�X�e���E�t�@�C��							System		Get/Set
' =			8  �i0b00001000�j	�f�B�X�N�h���C�u�E�{�����[���E���x��		Volume		Get
' =			16 �i0b00010000�j	�t�H���_�^�f�B���N�g��						Directory	Get
' =			32 �i0b00100000�j	�O��̃o�b�N�A�b�v�ȍ~�ɕύX����Ă����1	Archive		Get/Set
' =			64 �i0b01000000�j	�����N�^�V���[�g�J�b�g						Alias		Get
' =			128�i0b10000000�j	���k�t�@�C��								Compressed	Get
' ==================================================================
Public Function GetFileInfo( _
	ByVal sTrgtPath, _
	ByVal lGetInfoType, _
	ByRef vFileInfo _
)
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists( sTrgtPath ) Then
		'Do Nothing
	Else
		vFileInfo = ""
		GetFileInfo = False
		Exit Function
	End If
	
	Dim objFile
	Set objFile = objFSO.GetFile(sTrgtPath)
	
	vFileInfo = ""
	GetFileInfo = True
	Select Case lGetInfoType
		Case 1:		vFileInfo = objFile.Name				'�t�@�C����
		Case 2:		vFileInfo = objFile.Size				'�t�@�C���T�C�Y
		Case 3:		vFileInfo = objFile.Type				'�t�@�C�����
		Case 4:		vFileInfo = objFile.Drive				'�t�@�C���i�[��h���C�u
		Case 5:		vFileInfo = objFile.Path				'�t�@�C���p�X
		Case 6:		vFileInfo = objFile.ParentFolder		'�e�t�H���_
		Case 7:		vFileInfo = objFile.ShortName			'MS-DOS�`���t�@�C����
		Case 8:		vFileInfo = objFile.ShortPath			'MS-DOS�`���p�X
		Case 9:		vFileInfo = objFile.DateCreated			'�쐬����
		Case 10:	vFileInfo = objFile.DateLastAccessed	'�A�N�Z�X����
		Case 11:	vFileInfo = objFile.DateLastModified	'�X�V����
		Case 12:	vFileInfo = objFile.Attributes			'����
		Case Else:	GetFileInfo = False
	End Select
End Function
'	Call Test_GetFileInfo()
	Private Sub Test_GetFileInfo()
		Dim sBuf
		Dim bRet
		Dim vFileInfo
		sBuf = ""
		Dim sTrgtPath
		sTrgtPath = "C:\codes\vbs\_lib\FileSystem.vbs"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFileInfo( sTrgtPath,	1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C�����F"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	2, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���T�C�Y�F"		  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	3, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C����ށF"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	4, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���i�[��h���C�u�F" & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	5, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C���p�X�F"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	6, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�e�t�H���_�F"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	7, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS�`���t�@�C�����F"   & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	8, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	MS-DOS�`���p�X�F"		  & vFileInfo
		bRet = GetFileInfo( sTrgtPath,	9, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�쐬�����F"				  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 10, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�A�N�Z�X�����F"			  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 11, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�X�V�����F"				  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 12, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�����F"					  & vFileInfo
		bRet = GetFileInfo( sTrgtPath, 13, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�F"						  & vFileInfo
		sTrgtPath = "C:\codes\vbs\_lib\dummy.vbs"
		sBuf = sBuf & vbNewLine & sTrgtPath
		bRet = GetFileInfo( sTrgtPath,	1, vFileInfo) : sBuf = sBuf & vbNewLine & bRet & "	�t�@�C�����F"			  & vFileInfo
		MsgBox sBuf
	End Sub

