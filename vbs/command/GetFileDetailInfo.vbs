Option Explicit

'<<概要>>
'  ファイル情報を取得する。
'
'<<使用方法>>
'  GetFileDetailInfo.vbs <target_file_path> <info_type>...
'   ・ファイルパス（<target_file_path>）
'   ・ファイル情報種別（<info_type>）(※1)
'     (※1) ファイル情報種別
'         [引数]  [説明]                  [プロパティ名]      [データ型]              [Get/Set]   [出力例]
'         1       ファイル名              Name                vbString    文字列型    Get/Set     03 Ride Featuring Tony Matterhorn.MP3
'         2       ファイルサイズ          Size                vbLong      長整数型    Get         4286923
'         3       ファイル種類            Type                vbString    文字列型    Get         MPEG layer 3
'         4       ファイル格納先ドライブ  Drive               vbString    文字列型    Get         Z:
'         5       ファイルパス            Path                vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
'         6       親フォルダ              ParentFolder        vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice
'         7       MS-DOS形式ファイル名    ShortName           vbString    文字列型    Get         03 Ride Featuring Tony Matterhorn.MP3
'         8       MS-DOS形式パス          ShortPath           vbString    文字列型    Get         Z:\300_Musics\200_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3
'         9       作成日時                DateCreated         vbDate      日付型      Get         2015/08/19 0:54:45
'         10      アクセス日時            DateLastAccessed    vbDate      日付型      Get         2016/10/14 6:00:30
'         11      更新日時                DateLastModified    vbDate      日付型      Get         2016/10/14 6:00:30
'         12      属性                    Attributes          vbLong      長整数型    (※2)       32
'     (※2) 属性
'         [値]                [説明]                                      [属性名]    [Get/Set]
'         1  （0b00000001）   読み取り専用ファイル                        ReadOnly    Get/Set
'         2  （0b00000010）   隠しファイル                                Hidden      Get/Set
'         4  （0b00000100）   システム・ファイル                          System      Get/Set
'         8  （0b00001000）   ディスクドライブ・ボリューム・ラベル        Volume      Get
'         16 （0b00010000）   フォルダ／ディレクトリ                      Directory   Get
'         32 （0b00100000）   前回のバックアップ以降に変更されていれば1   Archive     Get/Set
'         64 （0b01000000）   リンク／ショートカット                      Alias       Get
'         128（0b10000000）   圧縮ファイル                                Compressed  Get

'===============================================================================
'= インクルード
'===============================================================================
Call Include( "%MYDIRPATH_CODES%\vbs\_lib\FileSystem.vbs" )  'GetFileInfo()

'===============================================================================
'= 設定値
'===============================================================================
Const bEXEC_TEST = False 'テスト用
Const sSCRIPT_NAME = "ファイル情報取得"

'===============================================================================
'= 本処理
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
'= メイン関数
'===============================================================================
Public Sub Main()
    If cArgs.Count >= 2 then
        Dim sTrgtPath
        sTrgtPath = cArgs(0)
        Dim cInfoTypes
        Set cInfoTypes = CreateObject("System.Collections.ArrayList")
        Dim lArgIdx
        For lArgIdx = 1 To (cArgs.Count - 1)
            If IsNumeric(cArgs(lArgIdx)) Then
                cInfoTypes.Add cArgs(lArgIdx)
            Else
                WScript.Echo "[error] <info_type> must be a number."
                WScript.Echo "  usage : GetFileDetailInfo.vbs <target_file_path> <info_type>..."
                Exit Sub
            End If
        Next
    Else
        WScript.Echo "[error] wrong number of argments"
        WScript.Echo "  usage : GetFileDetailInfo.vbs <target_file_path> <info_type>..."
        Exit Sub
    End If
    
    Dim vFileInfo
    Dim sOutputStr
    sOutputStr = sTrgtPath
    Dim sInfoType
    For Each sInfoType In cInfoTypes
        Dim bResult
        bResult = GetFileInfo( sTrgtPath, CLng(sInfoType), vFileInfo)
        If bResult = True Then
            sOutputStr = sOutputStr & vbTab & vFileInfo
        Else
            WScript.Echo "[error] GetFileInfo() failed."
            Exit Sub
        End If
    Next
    WScript.Echo sOutputStr
End Sub

'===============================================================================
'= 内部関数
'===============================================================================

'===============================================================================
'= テスト関数
'===============================================================================
Private Sub Test_Main() '{{{
    Const lTESTCASE_STRT = 1
    Const lTESTCASE_LAST = 6
    Dim lIdx
    For lIdx = lTESTCASE_STRT To lTESTCASE_LAST
        Dim sTestFuncName
        sTestFuncName = _
            "Test_Case" & _
            String(3 - Len(CStr(lIdx)), "0") & _
            CStr(lIdx)
        cArgs.Clear
        Dim oFuncPtr
        Set oFuncPtr = GetRef(sTestFuncName)
        WScript.Echo "=== " & sTestFuncName & " ==="
        oFuncPtr()
    Next
End Sub
Private Sub Test_Case001()
    
    cArgs.Add WScript.ScriptFullName
    cArgs.Add 1 'ファイル名
    cArgs.Add 2 'ファイルサイズ
    Call Main()
    
End Sub
Private Sub Test_Case002()
    cArgs.Add WScript.ScriptFullName
    cArgs.Add 1
    cArgs.Add 1
    cArgs.Add 1
    cArgs.Add 1
    Call Main()
End Sub
Private Sub Test_Case003()
    cArgs.Add WScript.ScriptFullName
    Call Main()
End Sub
Private Sub Test_Case004()
    cArgs.Add WScript.ScriptFullName
    cArgs.Add "aaa"
    Call Main()
End Sub
Private Sub Test_Case005()
    cArgs.Add WScript.ScriptFullName
    cArgs.Add 13
    Call Main()
End Sub
Private Sub Test_Case006()
    cArgs.Add "aaa"
    cArgs.Add 1
    Call Main()
End Sub
'}}}

'===============================================================================
'= インクルード関数
'===============================================================================
Private Function Include( ByVal sOpenFile ) '{{{
    sOpenFile = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings(sOpenFile)
    With CreateObject("Scripting.FileSystemObject").OpenTextFile( sOpenFile )
        ExecuteGlobal .ReadAll()
        .Close
    End With
End Function '}}}

