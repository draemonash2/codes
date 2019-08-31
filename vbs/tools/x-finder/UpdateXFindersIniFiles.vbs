Option Explicit

'==========================================================
'= インクルード
'==========================================================
Dim sMyDirPath
sMyDirPath = Replace( WScript.ScriptFullName, "\" & WScript.ScriptName, "" )
Call Include( "C:\codes\vbs\_lib\Log.vbs" )         'class LogMng
Call Include( "C:\codes\vbs\_lib\X-Finder.vbs" )    'UpdateIniFile()
Call Include( "C:\codes\vbs\_lib\FileSystem.vbs" )  'GetFileList()
                                                    'GetFileOrFolder()

'==========================================================
'= 本処理
'==========================================================
Dim objFSO  'FileSystemObjectの格納先
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

Dim sLogFilePath
Dim sRootDirPath
Dim sShortcutDirPath
Dim sIniRootDirPath
sLogFilePath = sMyDirPath & "\" & objFSO.GetBaseName( WScript.ScriptName ) & ".log"
sRootDirPath = sMyDirPath & "\.."
sShortcutDirPath = sRootDirPath & "\favorite"
sIniRootDirPath = sRootDirPath & "\data"

sLogFilePath = objFSO.GetAbsolutePathName( sLogFilePath )
sRootDirPath = objFSO.GetAbsolutePathName( sRootDirPath )
sShortcutDirPath = objFSO.GetAbsolutePathName( sShortcutDirPath )
sIniRootDirPath = objFSO.GetAbsolutePathName( sIniRootDirPath )

Dim oLogMng
Set oLogMng = New LogMng
Call oLogMng.Open( sLogFilePath, "w" )

oLogMng.Puts "sLogFilePath     : " & sLogFilePath
oLogMng.Puts "sRootDirPath     : " & sRootDirPath
oLogMng.Puts "sShortcutDirPath : " & sShortcutDirPath
oLogMng.Puts "sIniRootDirPath  : " & sIniRootDirPath
oLogMng.Puts ""

'iniファイル全削除
oLogMng.Puts "*** delete ini files ***"
Dim objFile
For Each objFile In objFSO.GetFolder( sIniRootDirPath ).Files
    If objFile.Name = "_favorite_data.ini" Then
        'Do Nothing
    ElseIf InStr( objFile.Name, ".ini" ) Then
        oLogMng.Puts objFile.Path
        objFSO.DeleteFile objFile.Path, True
    Else
        'Do Nothing
    End If
Next
oLogMng.Puts ""

'ショートカット ファイル/フォルダ一覧取得
Dim asFileList()
ReDim Preserve asFileList(-1)
Call GetFileList( sShortcutDirPath, asFileList, 0 )

Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

'Ini ファイル作成
oLogMng.Puts "*** create ini files ***"
Dim sFileDirPath
Dim sFileDirParentDirPath
Dim sIniFileName
Dim sIniFilePath
Dim sIniTrgtFileName
Dim sItemName
Dim sItemPath
Dim sItemType
Dim sItemIcon
Dim sItemExt
Dim lIdx
For lIdx = 0 to UBound( asFileList )
    sFileDirPath = asFileList( lIdx )
    If sFileDirPath = sShortcutDirPath Then
        'Do Nothing
    Else
        If GetFileOrFolder( sFileDirPath ) = 1 Then 'ファイル
            If objFSO.GetExtensionName( sFileDirPath ) = "lnk" Then
                sFileDirParentDirPath       = objFSO.GetParentFolderName( sFileDirPath )
                sIniFileName                = "_" & Replace( Replace( sFileDirParentDirPath, sRootDirPath & "\", "" ), "\", "_" ) & ".ini"
                sIniFilePath                = sIniRootDirPath & "\" & sIniFileName
                sItemName                   = objFSO.GetBaseName( sFileDirPath )
                sItemPath                   = """" & objWshShell.CreateShortcut( sFileDirPath ).TargetPath & """"
                sItemType                   = 1
                sItemIcon                   = ""
                sItemExt                    = ""
                Call UpdateIniFile( sIniFilePath, sItemName, sItemPath, sItemType, sItemIcon, sItemExt )
                Call oLogMng.Puts( "file   : " & chr(9) & sIniFilePath & chr(9) & sItemName & chr(9) & sItemPath & chr(9) & sItemType & chr(9) & sItemIcon & chr(9) & sItemExt )
            Else
                'Do Nothing
            End If
        ElseIf GetFileOrFolder( sFileDirPath ) = 2 Then 'フォルダ
            sFileDirParentDirPath           = objFSO.GetParentFolderName( sFileDirPath )
            sIniFileName                    = "_" & Replace( Replace( sFileDirParentDirPath, sRootDirPath & "\", "" ), "\", "_" ) & ".ini"
            sIniFilePath                    = sIniRootDirPath & "\" & sIniFileName
            sIniTrgtFileName                = "_" & Replace( Replace( sFileDirPath, sRootDirPath & "\", "" ), "\", "_" ) & ".ini"
            sItemName                       = objFSO.GetFolder( sFileDirPath ).Name
            sItemPath                       = "Extra:" & sIniTrgtFileName
            sItemType                       = 1
            sItemIcon                       = "shell32.dll,3"
            sItemExt                        = ""
            Call UpdateIniFile( sIniFilePath, sItemName, sItemPath, sItemType, sItemIcon, sItemExt )
            Call oLogMng.Puts( "folder : " & chr(9) & sIniFilePath & chr(9) & sItemName & chr(9) & sItemPath & chr(9) & sItemType & chr(9) & sItemIcon & chr(9) & sItemExt )
        Else
            Dim sLogMsg
            sLogMsg = "[error  ] target path is invalid! " & sFileDirPath
            oLogMng.Puts sLogMsg
            MsgBox sLogMsg
            Wscript.Quit()
        End If
    End If
Next

Call oLogMng.Close()
Set oLogMng = Nothing

'==========================================================
'= 関数定義
'==========================================================
' 外部プログラム インクルード関数
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

