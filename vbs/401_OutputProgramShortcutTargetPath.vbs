Option Explicit

' �T�v�j
'�w�肳�ꂽ�t�H���_�z���ɑ��݂���V���[�g�J�b�g(*.lnk)�̎w������A
'�ꗗ�����ďo�͂���B
'
'���s���@�j
'1. �{�X�N���v�g�t�@�C�������s����B

'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\_lib\FileSystem.vbs" )
Call Include( sMyDirPath & "\_lib\Log.vbs" )

'==========================================================
'= �ݒ�l
'==========================================================

'==========================================================
'= �{����
'==========================================================
MsgBox "�{�v���O�����͊Ǘ��Ҍ������K�v�ƂȂ�ꍇ������܂��B" & vbNewLine & "�G���[�����������ꍇ�A�Ǘ��Ҍ����ɂĎ��s���Ă��������B"

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
Dim sTrgtDir
sTrgtDir = objWshShell.CurrentDirectory
'sTrgtDir = objWshShell.SpecialFolders("StartMenu")
Dim bIsContinue
Do
    Dim vAnswer
    vAnswer = MsgBox( "�ȉ���ΏۂɎ��s���܂��B���s���܂����H" & vbNewLine & sTrgtDir, vbOkCancel )
    If vAnswer = vbOk Then
        bIsContinue = False
    Else
        vAnswer = MsgBox( "�����𑱂��܂����H", vbOkCancel )
        If vAnswer = vbCancel Then
            WScript.Quit
        Else
            sTrgtDir = InputBox ( "�Ώۃf�B���N�g�����w�肵�Ă��������B" )
            bIsContinue = True
        End If
    End If
Loop While bIsContinue = True

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
oLogMng.Puts( "" )
oLogMng.Puts( "### Result ###" )
oLogMng.Puts( "<Type>" & chr(9) & "<sFileDirPath>" & chr(9) & "<sTargetPath>" )

Dim i
For i = 0 to UBound( asFileList ) - 1
    Dim sFileDirPath
    sFileDirPath = asFileList(i)
    
    If objFSO.GetExtensionName( sFileDirPath ) = "lnk" Then
        With objWshShell.CreateShortcut( sFileDirPath )
            oLogMng.Puts( "[ShrtCt  ]" & chr(9) & sFileDirPath & chr(9) & .TargetPath )
        End With
    Else
        oLogMng.Puts( "[NoShrtCt]" & chr(9) & sFileDirPath )
    End If
Next

oLogMng.Close()
Set oLogMng = Nothing

MsgBox _
    "�ȉ��Ƀv���O�����V���[�g�J�b�g�̎w������o�͂��܂����B" & vbNewLine & _
    "  " & sLogFilePath

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

