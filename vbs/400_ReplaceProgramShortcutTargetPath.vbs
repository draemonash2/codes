Option Explicit

'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\lib\FileSystem.vbs" )
Call Include( sMyDirPath & "\lib\Log.vbs" )

'==========================================================
'= �ݒ�l
'==========================================================
Const EXE_PATH_ORG = "C:\Users\draem_000\Documents\Amazon Drive\100_Programs\program\prg_exe"
Const EXE_PATH_NEW = "C:\prg_exe"

'==========================================================
'= �{����
'==========================================================
MsgBox "�{�v���O�����͊Ǘ��Ҍ������K�v�ƂȂ�ꍇ������܂��B" & vbNewLine & "�G���[�����������ꍇ�A�Ǘ��Ҍ����ɂĎ��s���Ă��������B"

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim sTrgtDir
If WScript.Arguments.Count = 0 Then
'	sTrgtDir = objWshShell.SpecialFolders("StartMenu")
	sTrgtDir = InputBox ( "�Ώۃf�B���N�g�����w�肵�Ă��������B" )
Else
	sTrgtDir = WScript.Arguments(0)
End If

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim sMyFileBaseName
sMyFileBaseName = objFSO.GetBaseName( WScript.ScriptFullName )

Dim oLogMng
Set oLogMng = New LogMng
Dim sLogFilePath
sLogFilePath = sTrgtDir & "\" & sMyFileBaseName & ".log"
Call oLogMng.Open( sLogFilePath, "w" )

Dim asFileList
Call GetFileList2( sTrgtDir, asFileList, 1 )

oLogMng.Puts( "target directory path : " & sTrgtDir )
oLogMng.Puts( "org path              : " & EXE_PATH_ORG )
oLogMng.Puts( "new path              : " & EXE_PATH_NEW )
oLogMng.Puts( "" )
oLogMng.Puts( "[Result]" & chr(9) & "[sFileDirPath]" & chr(9) & "[sOrgDirPath]" & chr(9) & "[sNewDirPath]" )

Dim i
For i = 0 to UBound( asFileList ) - 1
	Dim sFileDirPath
	sFileDirPath = asFileList(i)
	
	Dim sOrgDirPath
	Dim sNewDirPath
	If objFSO.GetExtensionName( sFileDirPath ) = "lnk" Then
		With objWshShell.CreateShortcut( sFileDirPath )
			sOrgDirPath = .TargetPath
			If InStr( sOrgDirPath, EXE_PATH_ORG ) > 0 Then
				sNewDirPath = Replace( sOrgDirPath, EXE_PATH_ORG, EXE_PATH_NEW )
					.TargetPath = sNewDirPath
					.Save
				oLogMng.Puts( "[Replaced]" & chr(9) & sFileDirPath & chr(9) & sOrgDirPath & chr(9) & sNewDirPath )
			Else
				oLogMng.Puts( "[UnMatch ]" & chr(9) & sFileDirPath & chr(9) & sOrgDirPath )
			End If
		End With
	Else
		oLogMng.Puts( "[NoShrtCt]" & chr(9) & sFileDirPath )
	End If
Next

oLogMng.Close()
Set oLogMng = Nothing

MsgBox _
	"�v���O�����̃V���[�g�J�b�g�̎w�����u�����܂����B" & vbNewLine & _
	"���F" & EXE_PATH_ORG & vbNewLine & _
	"��F" & EXE_PATH_NEW & vbNewLine & _
	"" & vbNewLine & _
	"���u" & sLogFilePath & "�v�ɒu�����ʂ��o�͂��Ă��܂�"

'==========================================================
'= �֐���`
'==========================================================
' �O���v���O���� �C���N���[�h�֐�
Function Include( _
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

