Option Explicit

'TODO:�v�R�}���h��

' �T�v�j
'�w�肳�ꂽ�t�H���_�z���ɑ��݂���V���[�g�J�b�g(*.lnk)�̎w������A�u������B
'
'���s���@�j
'1. �ݒ�l���C������B
'2. �{�X�N���v�g�t�@�C�������s����B

'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetFileList2()
Call Include( "C:\codes\vbs\_lib\Log.vbs" )         'class LogMng

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
oLogMng.Puts( "org path              : " & EXE_PATH_ORG )
oLogMng.Puts( "new path              : " & EXE_PATH_NEW )
oLogMng.Puts( "" )
oLogMng.Puts( "### Legend ###" )
oLogMng.Puts( "  Replaced : replaced target pathes" )
oLogMng.Puts( "  UnMatch  : replace word is nothing at target path" )
oLogMng.Puts( "  NoShrtCt : not a program shortcut file" )
oLogMng.Puts( "### Result ###" )
oLogMng.Puts( "<Result>" & chr(9) & "<sFileDirPath>" & chr(9) & "<sOrgDirPath>" & chr(9) & "<sNewDirPath>" )

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
    "�ȉ��Ƀv���O�����V���[�g�J�b�g�̎w����u�����ʂ��o�͂��܂����B" & vbNewLine & _
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

