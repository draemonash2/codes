'==========================================================
'= �C���N���[�h
'==========================================================
Dim objWshShell
Set objWshShell = WScript.CreateObject( "WScript.Shell" )
Call Include( objWshShell.CurrentDirectory & "\String.vbs" )

'==========================================================
'= �{����
'==========================================================
'�������ɏ�����������

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
