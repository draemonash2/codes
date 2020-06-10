'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorer������s�A1:X-Finder������s�Aother:�f�o�b�O���s

'####################################################################
'### �ݒ�
'####################################################################
Const TEMP_FILE_NAME = "diff_target_path.tmp"

'####################################################################
'### �{����
'####################################################################
Const PROG_NAME = "WinMerge�Ŕ�r"

Dim bIsContinue
bIsContinue = True

Dim objWshShell
Dim objFSO
Dim sExePath
Dim cSelected
Dim sTmpFileDeleteMode

Set objWshShell = WScript.CreateObject("WScript.Shell")

'*** �I���t�@�C���擾 ***
If bIsContinue = True Then
    If EXECUTION_MODE = 0 Then 'Explorer������s
        Set cSelected = CreateObject("System.Collections.ArrayList")
        Dim sArg
        If WScript.Arguments.Count = 0 Then
            sTmpFileDeleteMode = True
        Else
            For Each sArg In WScript.Arguments
                cSelected.add sArg
            Next
            sTmpFileDeleteMode = False
        End If
    ElseIf EXECUTION_MODE = 1 Then 'X-Finder������s
        Set cSelected = WScript.Col(WScript.Env("Selected"))
        If TMP_FILE_DELETE_MODE = 1 Then
            sTmpFileDeleteMode = True
        Else
            sTmpFileDeleteMode = False
        End If
    Else
        MsgBox "�f�o�b�O���[�h�ł��B"
        Set cSelected = CreateObject("System.Collections.ArrayList")
        cSelected.Add "C:\prg_exe\X-Finder\script\FileNameCopy.vbs"
        cSelected.Add "C:\prg_exe\X-Finder\script\FilePathCopy.vbs"
        sTmpFileDeleteMode = False
    End If
Else
    'Do Nothing
End If

'*** ��r ***
If bIsContinue = True Then
    sTmpPath = "C:\Users\"& CreateObject("WScript.Network").UserName &"\AppData\Local\Temp\" & TEMP_FILE_NAME
    If sTmpFileDeleteMode = True Then
        If objFSO.FileExists( sTmpPath ) Then
            objFSO.DeleteFile sTmpPath, True
        Else
            'Do Nothing
        End If
    Else
        Set objWshShell = WScript.CreateObject("WScript.Shell")
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        sExePath = objWshShell.Environment("System").Item("MYSYSPATH_WINMERGE")
        If sExePath = "" then
            MsgBox "���ϐ����ݒ肳��Ă��܂���B" & vbNewLine & "�����𒆒f���܂��B", vbYes, PROG_NAME
            WScript.Quit
        end if
        
        Dim sExecCmd
        Dim sDiffPath1
        Dim sDiffPath2
        Dim objTxtFile
        If cSelected.Count >= 3 Then
            sExecCmd = """" & sExePath & """ -r """ & cSelected.Item(0) & """ """ & cSelected.Item(1) & """ """ & cSelected.Item(2) & """"
            objWshShell.Run sExecCmd, 3, False
        ElseIf cSelected.Count = 2 Then
            sExecCmd = """" & sExePath & """ -r """ & cSelected.Item(0) & """ """ & cSelected.Item(1) & """"
            objWshShell.Run sExecCmd, 3, False
        ElseIf cSelected.Count = 1 Then
            sDiffPath1 = cSelected.Item(0)
            If  objFSO.FileExists( sTmpPath ) Then
                Set objTxtFile = objFSO.OpenTextFile( sTmpPath, 1 )
                sDiffPath2 = objTxtFile.ReadLine
                objTxtFile.Close
                Set objTxtFile = Nothing
                sExecCmd = """" & sExePath & """ -r """ & sDiffPath2 & """ """ & sDiffPath1 & """"
                objWshShell.Run sExecCmd, 3, False
                objFSO.DeleteFile sTmpPath, True
            Else
                Set objTxtFile = objFSO.OpenTextFile( sTmpPath, 2, True )
                objTxtFile.WriteLine sDiffPath1
                objTxtFile.Close
                Set objTxtFile = Nothing
                MsgBox "�ȉ����r�ΏۂƂ��đI�����܂��B" & vbNewLine & vbNewLine & sDiffPath1
            End If
        Else
            MsgBox "�t�@�C�����I������Ă��܂���", vbYes, PROG_NAME
        End If
    End If
Else
    'Do Nothing
End If
