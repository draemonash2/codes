Option Explicit

'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\lib\String.vbs" )

'==========================================================
'= �{����
'==========================================================
Const INDEX_MAX = 500
'Const lContextLenBMax = 40

If WScript.Arguments.Count = 0 then
    MsgBox "�����o�͂������t�@�C����{�X�N���v�g�Ƀh���b�O���h���b�v���Ă��������B"
    MsgBox "�v���O�����𒆒f���܂��B"
    WScript.Quit(-1)
Else
    Dim lArgIdx
    For lArgIdx = 0 to WScript.Arguments.Count - 1
        Dim sDirPath
        Dim sFileName
        Dim sFilePath
        sFilePath = WScript.Arguments( lArgIdx )
        sDirPath = GetDirPath( sFilePath )
        sFileName = GetFileName( sFilePath )
        
        Dim objFolder
        Dim objFile
        Set objFolder = CreateObject( "Shell.Application" ).Namespace( sDirPath )
        Set objFile = objFolder.ParseName( sFileName )
        
        Dim sLogPath
        sLogPath = sDirPath & "\" & Replace(sFileName, ".", "_") & ".txt"
'        sLogPath = sDirPath & "\" & Replace(Replace(sFileName," ", "_"), ".", "_") & ".txt"
        
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        On Error Resume Next
        Dim objLogFile
        Set objLogFile = objFSO.OpenTextFile( sLogPath, 2, True )
        If Err.Number <> 0 Then
            MsgBox Err.Number & "�F" & Err.Description & vbNewLine & _
                   sLogPath
            WScript.Quit
        End If
        On Error Goto 0
        
        'MsgBox "�w�肳�ꂽ�t�@�C���̃t�@�C�������ȉ��ɏo�͂��܂��B" & vbNewLine & _
        '      "  [�t�@�C���p�X] " & sLogPath & vbNewLine & _
        '       "  [�����R�[�h] Unicode"
        
        Dim sItem
        Dim sContext
        
        '*** ���ڐ������ڕ������Z�o ***
        Dim lContextLenBMax
        Dim lIdx
        lContextLenBMax = 0
        For lIdx = 0 to INDEX_MAX
            sContext = objFolder.GetDetailsOf( objFolder.Items, lIdx )
            If sContext = "" Then
                'Do Nothing
            Else
                If Len( sContext ) > lContextLenBMax Then
                    lContextLenBMax = LenByte( sContext )
                Else
                    'Do Nothing
                End If
            End If
        Next
        
        '*** ���ڏo�� ***
        objLogFile.WriteLine "+-----+-" & String( lContextLenBMax , "-" ) & "-+-------------------------------------------------------"
        objLogFile.WriteLine "| idx | ���ږ�" & String( lContextLenBMax + 1 - LenByte("���ږ�"), " " ) & "| �l"
        objLogFile.WriteLine "+-----+-" & String( lContextLenBMax , "-" ) & "-+-------------------------------------------------------"
        
        Dim lContextNum
        lContextNum = 0
        For lIdx = 0 to INDEX_MAX
            sContext = objFolder.GetDetailsOf( objFolder.Items, lIdx )
            sItem = objFolder.GetDetailsOf( objFile, lIdx )
            
            If sContext = "" Or sItem = "" Then
                'Do Nothing
            Else
                On Error Resume Next
                Do
                    objLogFile.WriteLine "| " & String( 3 - Len(lIdx), " " ) & lIdx & " | " & _
                                          sContext & String( lContextLenBMax - LenByte(sContext), " " ) & " | " & _
                                          sItem
                    If Err.Number <> 0 Then
                        sItem = Right( sItem, Len(sItem) - 1 )
                        Err.Clear
                    Else
                        Exit Do
                    End If
                Loop While True
                On Error Goto 0
                lContextNum = lContextNum + 1
            End If
        Next
        objLogFile.WriteLine "+-----+-" & String( lContextLenBMax , "-" ) & "-+-------------------------------------------------------"
        objLogFile.WriteLine "�y���ڐ��z" & lContextNum
        objLogFile.Close
        
        Set objFolder = Nothing
        Set objFile = Nothing
        Set objFSO = Nothing
        Set objLogFile = Nothing
    Next
    MsgBox "�����I"
End if

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
    On Error Resume Next
    Set objVbsFile = objFSO.OpenTextFile( sOpenFile )
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbNewLine & _
               sOpenFile
        WScript.Quit
    End If
    On Error Goto 0
    
    ExecuteGlobal objVbsFile.ReadAll()
    objVbsFile.Close
    
    Set objVbsFile = Nothing
    Set objFSO = Nothing
End Function
