'===============================================================================
'= �C���N���[�h
'===============================================================================
Call Include( "C:\codes\vbs\_lib\String.vbs" ) '��()

'===============================================================================
'= �{����
'===============================================================================
'�������ɏ�����������

'===============================================================================
'= �C���N���[�h�֐�
'===============================================================================
Private Function Include( _
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
