Option Explicit

'==============================================================================
'�y�T�v�z
'	�t�@�C��/�t�H���_���R�s�[����B
'	�ړ���̃t�H���_�����݂��Ȃ��ꍇ�A�t�H���_���쐬���Ă���R�s�[����B
'
'�y�g�p���@�z
'	copy_to_dir.vbs <source_path> <destination_path>
'
'�y�g�p��z
'	1) copy_to_dir.vbs c:\codes\vbs\test.txt c:\test\test.txt
'	2) copy_to_dir.vbs c:\codes\vbs c:\test\vbs
'		c:\codes\vbs
'			�� a.txt
'			�� b
'				�� c.txt
'		��
'		c:\test\vbs
'			�� a.txt
'			�� b
'				�� c.txt
'
'�y�o�������z
'	�Ȃ�
'
'�y���������z
'	1.0.0	2019/05/12	�V�K�쐬
'==============================================================================

'==============================================================================
' �ݒ�
'==============================================================================

'==============================================================================
' �{����
'==============================================================================
'�����`�F�b�N
If WScript.Arguments.Count = 2 Then
	'Do Nothing
Else
	Wscript.quit
End If

dim sSrcPath
dim sDstPath
sSrcPath = Replace(WScript.Arguments(0), "/", "\")
sDstPath = Replace(WScript.Arguments(1), "/", "\")

Dim lSrcPathType
lSrcPathType = GetFileOrFolder(sSrcPath)

dim sDstParDir
sDstParDir = GetDirPath( sDstPath )

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If lSrcPathType = 1 Then '�t�@�C��
	call CreateDirectry( sDstParDir )
	objFSO.CopyFile sSrcPath, sDstPath
ElseIf lSrcPathType = 2 Then '�t�H���_
	call CreateDirectry( sDstParDir )
	objFSO.CopyFolder sSrcPath, sDstPath
Else '������
'	WScript.Echo "�t�@�C�������݂��܂���"
End If

Set objFSO = Nothing

'==============================================================================
' ���C�u����
'==============================================================================
' ==================================================================
' = �T�v	�t�H���_���쐬����
' = ����	sDirPath		String		[in]	�쐬�Ώۃt�H���_
' = �ߒl	�Ȃ�
' = �o��	�E�쐬�Ώۃt�H���_�̐e�f�B���N�g�������݂��Ȃ��ꍇ�A
' =			  �ċA�I�ɐe�t�H���_���쐬����
' =			�E�t�H���_�����ɑ��݂��Ă���ꍇ�͉������Ȃ�
' ==================================================================
Public Function CreateDirectry( _
	ByVal sDirPath _
)
	Dim sParentDir
	Dim oFileSys
	
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	
	sParentDir = oFileSys.GetParentFolderName(sDirPath)
	
	'�e�f�B���N�g�������݂��Ȃ��ꍇ�A�ċA�Ăяo��
	If oFileSys.FolderExists( sParentDir ) = False Then
		Call CreateDirectry( sParentDir )
	End If
	
	'�f�B���N�g���쐬
	If oFileSys.FolderExists( sDirPath ) = False Then
		oFileSys.CreateFolder sDirPath
	End If
	
	Set oFileSys = Nothing
End Function

' ==================================================================
' = �T�v	�t�@�C�����t�H���_���𔻒肷��
' = ����	sChkTrgtPath	String		[in]	�`�F�b�N�Ώۃt�H���_
' = �ߒl					Long				���茋��
' =													1) �t�@�C��
' =													2) �t�H���_�[
' =													0) �G���[�i���݂��Ȃ��p�X�j
' = �o��	FileSystemObject ���g���Ă���̂ŁA�t�@�C��/�t�H���_��
' =			���݊m�F�ɂ��g�p�\�B
' ==================================================================
Public Function GetFileOrFolder( _
	ByVal sChkTrgtPath _
)
	Dim oFileSys
	Dim bFolderExists
	Dim bFileExists
	
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	bFolderExists = oFileSys.FolderExists(sChkTrgtPath)
	bFileExists = oFileSys.FileExists(sChkTrgtPath)
	Set oFileSys = Nothing
	
	If bFolderExists = False And bFileExists = True Then
		GetFileOrFolder = 1 '�t�@�C��
	ElseIf bFolderExists = True And bFileExists = False Then
		GetFileOrFolder = 2 '�t�H���_�[
	Else
		GetFileOrFolder = 0 '�G���[�i���݂��Ȃ��p�X�j
	End If
End Function

' ==================================================================
' = �T�v	�w�肳�ꂽ�t�@�C���p�X����t�H���_�p�X�𒊏o����
' = ����	sFilePath	String	[in]  �t�@�C���p�X
' = �ߒl				String		  �t�H���_�p�X
' = �o��	���[�J���t�@�C���p�X�i��Fc:\test�j�� URL �i��Fhttps://test�j
' =			���w��\
' ==================================================================
Public Function GetDirPath( _
	ByVal sFilePath _
)
	If InStr( sFilePath, "\" ) Then
		GetDirPath = RemoveTailWord( sFilePath, "\" )
	ElseIf InStr( sFilePath, "/" ) Then
		GetDirPath = RemoveTailWord( sFilePath, "/" )
	Else
		GetDirPath = sFilePath
	End If
End Function
	'Call Test_GetDirPath()
	Private Sub Test_GetDirPath()
		Dim Result
		Result = "[Result]"
		Result = Result & vbNewLine & GetDirPath( "C:\test\a.txt" )    ' C:\test
		Result = Result & vbNewLine & GetDirPath( "http://test/a" )    ' http://test
		Result = Result & vbNewLine & GetDirPath( "C:_test_a.txt" )    ' C:_test_a.txt
		MsgBox Result
	End Sub

' ==================================================================
' = �T�v	������؂蕶���ȍ~�̕��������������B
' = ����	sStr		String	[in]  �������镶����
' = ����	sDlmtr		String	[in]  ��؂蕶��
' = �ߒl				String		  ����������
' = �o��	�Ȃ�
' ==================================================================
Public Function RemoveTailWord( _
	ByVal sStr, _
	ByVal sDlmtr _
)
	Dim sTailWord
	Dim lRemoveLen
	
	If sStr = "" Then
		RemoveTailWord = ""
	Else
		If sDlmtr = "" Then
			RemoveTailWord = sStr
		Else
			If InStr(sStr, sDlmtr) = 0 Then
				RemoveTailWord = sStr
			Else
				sTailWord = ExtractTailWord(sStr, sDlmtr)
				lRemoveLen = Len(sDlmtr) + Len(sTailWord)
				RemoveTailWord = Left(sStr, Len(sStr) - lRemoveLen)
			End If
		End If
	End If
End Function
	'Call Test_RemoveTailWord()
	Private Sub Test_RemoveTailWord()
		Dim Result
		Result = "[Result]"
		Result = Result & vbNewLine & "*** test start! ***"
		Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "\" )	' C:\test
		Result = Result & vbNewLine & RemoveTailWord( "C:\test\a", "\" )		' C:\test
		Result = Result & vbNewLine & RemoveTailWord( "C:\test\", "\" )			' C:\test
		Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\" )			' C:
		Result = Result & vbNewLine & RemoveTailWord( "C:\test", "\\" )			' C:\test
		Result = Result & vbNewLine & RemoveTailWord( "", "\" )					' 
		Result = Result & vbNewLine & RemoveTailWord( "a.txt", "\" )			' a.txt�i�t�@�C�������ǂ����͔��f���Ȃ��j
		Result = Result & vbNewLine & RemoveTailWord( "C:\test\a.txt", "" )		' C:\test\a.txt
		Result = Result & vbNewLine & "*** test finished! ***"
		MsgBox Result
	End Sub

' ==================================================================
' = �T�v	������؂蕶���ȍ~�̕������ԋp����B
' = ����	sStr		String	[in]  �������镶����
' = ����	sDlmtr		String	[in]  ��؂蕶��
' = �ߒl				String		  ���o������
' = �o��	�Ȃ�
' ==================================================================
Public Function ExtractTailWord( _
	ByVal sStr, _
	ByVal sDlmtr _
)
	Dim asSplitWord
	
	If Len(sStr) = 0 Then
		ExtractTailWord = ""
	Else
		ExtractTailWord = ""
		asSplitWord = Split(sStr, sDlmtr)
		ExtractTailWord = asSplitWord(UBound(asSplitWord))
	End If
End Function
	'Call Test_ExtractTailWord()
	Private Sub Test_ExtractTailWord()
		Dim Result
		Result = "[Result]"
		Result = Result & vbNewLine & "*** test start! ***"
		Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "\" )	' a.txt
		Result = Result & vbNewLine & ExtractTailWord( "C:\test\a", "\" )		' a
		Result = Result & vbNewLine & ExtractTailWord( "C:\test\", "\" )		' 
		Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\" )			' test
		Result = Result & vbNewLine & ExtractTailWord( "C:\test", "\\" )		' C:\test
		Result = Result & vbNewLine & ExtractTailWord( "a.txt", "\" )			' a.txt
		Result = Result & vbNewLine & ExtractTailWord( "", "\" )				' 
		Result = Result & vbNewLine & ExtractTailWord( "C:\test\a.txt", "" )	' C:\test\a.txt
		Result = Result & vbNewLine & "*** test finished! ***"
		MsgBox Result
	End Sub
