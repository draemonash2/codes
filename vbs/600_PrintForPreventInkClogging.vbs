'==========================================================
'= �C���N���[�h
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( sMyDirPath & "\_lib\Excel.vbs" )

'==========================================================
'= �{����
'==========================================================
Call PrintExcelSheet( _
    "C:\Users\draem_000\Documents\Dropbox\100_Documents\111_�y�����z���d�����i��PC���g�ѓd�b\PC\�v�����^�C���N�l�܂�h�~�p����y�[�W.xlsx", _
    "Sheet1", _
    1, _
    1, _
    1 _
)

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
