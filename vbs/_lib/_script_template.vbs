Option Explicit

'<<�T�v>>
'  
'  
'<<�g�p���@>>
'  
'  
'<<�d�l>>
'  �E

'===============================================================================
'= �C���N���[�h
'===============================================================================
'Call Include( "%MYDIRPATH_CODES%\vbs\_lib\String.vbs" )     'ConvDate2String()

'===============================================================================
'= �ݒ�l
'===============================================================================
Const bEXEC_TEST = False '�e�X�g�p
Const sSCRIPT_NAME = ��

'===============================================================================
'= �{����
'===============================================================================
Dim cArgs '{{{
Set cArgs = CreateObject("System.Collections.ArrayList")

If bEXEC_TEST = True Then
    Call Test_Main()
Else
    Dim vArg
    For Each vArg in WScript.Arguments
        cArgs.Add vArg
    Next
    Call Main()
End If '}}}

'===============================================================================
'= ���C���֐�
'===============================================================================
Public Sub Main()
    Dim sTrgtPath
    Dim lBakFileNumMax
    Dim sBakLogFilePath
    If cArgs.Count >= 1 Then
        sTrgtPath = cArgs(0)
    Else
        WScript.Echo "�������w�肵�Ă��������B�v���O�����𒆒f���܂��B"
        Exit Sub
    End If
End Sub

'===============================================================================
'= �����֐�
'===============================================================================

'===============================================================================
'= �e�X�g�֐�
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTestCase = 1
    
    cArgs.Add sTrgtFilePath
    
    MsgBox "=== test start ==="
    
    Select Case lTestCase
        Case 1
            Call Main()
            MsgBox "1 ���s��"
        Case Else
            Call Main()
    End Select
    
    MsgBox "=== test finished ==="
End Sub '}}}

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( ByVal sOpenFile ) '{{{
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function '}}}

