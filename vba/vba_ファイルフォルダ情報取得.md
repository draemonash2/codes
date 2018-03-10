# ファイル・フォルダ情報
## プロパティ一覧
- ファイル情報（File オブジェクト）

| プロパティ名     | 説明                   |>         | データ型        | Get/Set | 出力例                                                                                    |
|:---|:---|:---|:---|:---|:---|
| Name             | ファイル名             | vbString | 文字列型        | Get/Set | 03 Ride Featuring Tony Matterhorn.MP3                                                     |
| Size             | ファイルサイズ         | vbLong   | 長整数型 (Long) | Get     | 4286923                                                                                   |
| Type             | ファイル種類           | vbString | 文字列型        | Get     | MPEG layer 3                                                                              |
| Drive            | ファイル格納先ドライブ | vbString | 文字列型        | Get     | Z:                                                                                        |
| Path             | ファイルパス           | vbString | 文字列型        | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3 |
| ParentFolder     | 親フォルダ             | vbString | 文字列型        | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice                                       |
| ShortName        | MS-DOS形式ファイル名   | vbString | 文字列型        | Get     | 03 Ride Featuring Tony Matterhorn.MP3                                                     |
| ShortPath        | MS-DOS形式パス         | vbString | 文字列型        | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice\03 Ride Featuring Tony Matterhorn.MP3 |
| DateCreated      | 作成日時               | vbDate   | 日付型 (Date)   | Get     | 2015/08/19 0:54:45                                                                        |
| DateLastAccessed | アクセス日時           | vbDate   | 日付型 (Date)   | Get     | 2016/10/14 6:00:30                                                                        |
| DateLastModified | 更新日時               | vbDate   | 日付型 (Date)   | Get     | 2016/10/14 6:00:30                                                                        |
| Attributes       | 属性                   | vbLong   | 長整数型 (Long) | (※)    | 32                                                                                        |

- フォルダ情報（Folder オブジェクト）

| プロパティ名     | 説明                   |>          | データ型           | Get/Set | 出力例                                              |
|:---|:---|:---|:---|:---|:---|
| Name             | フォルダ名             | vbString  | 文字列型           | Get/Set | Sacrifice                                           |
| Size             | フォルダサイズ         | vbLong    | 長整数型 (Long)    | Get     | 80613775                                            |
| Type             | ファイル種類           | vbString  | 文字列型           | Get     | ファイル フォルダー                                 |
| Drive            | ファイル格納先ドライブ | vbString  | 文字列型           | Get     | Z:                                                  |
| Path             | フォルダパス           | vbString  | 文字列型           | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice |
| IsRootFolder     | ルート フォルダ        | vbBoolean | ブール型 (Boolean) | Get     | False                                               |
| ShortName        | MS-DOS形式ファイル名   | vbString  | 文字列型           | Get     | Sacrifice                                           |
| ShortPath        | MS-DOS形式パス         | vbString  | 文字列型           | Get     | Z:\300\_Musics\200\_DanceHall\Artist\Alaine\Sacrifice |
| DateCreated      | 作成日時               | vbDate    | 日付型 (Date)      | Get     | 2015/08/19 0:54:44                                  |
| DateLastAccessed | アクセス日時           | vbDate    | 日付型 (Date)      | Get     | 2015/08/19 0:54:44                                  |
| DateLastModified | 更新日時               | vbDate    | 日付型 (Date)      | Get     | 2015/04/18 3:38:36                                  |
| Attributes       | 属性                   | vbLong    | 長整数型 (Long)    | (※)    | 16                                                  |

- 属性（※）

| 属性名     | 説明                                      | Get/Set(※) | ビット            |
|:---|:---|:---|:---|
| ReadOnly   | 読み取り専用ファイル                      | Get/Set     | 1（0b00000001）   |
| Hidden     | 隠しファイル                              | Get/Set     | 2（0b00000010）   |
| System     | システム・ファイル                        | Get/Set     | 4（0b00000100）   |
| Volume     | ディスクドライブ・ボリューム・ラベル      | Get         | 8（0b00001000）   |
| Directory  | フォルダ／ディレクトリ                    | Get         | 16（0b00010000）  |
| Archive    | 前回のバックアップ以降に変更されていれば1 | Get/Set     | 32（0b00100000）  |
| Alias      | リンク／ショートカット                    | Get         | 64（0b01000000）  |
| Compressed | 圧縮ファイル                              | Get         | 128（0b10000000） |

## 実行例
``` vba
Sub test()
	Dim sDirPath As String
	Dim sFileName As String
	Dim sFilePath As String
	sDirPath = "Z:\300_Musics\200_Reggae@Jamaica\Artist\Alaine\Sacrifice"
	sFileName = "03 Ride Featuring Tony Matterhorn.MP3"
	sFilePath = sDirPath & "\" & sFileName
	 
	Dim objFSO As Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	'=====================================================
	' ファイル情報
	'=====================================================
	Dim objFile As Object
	Set objFile = objFSO.GetFile(sFilePath)
	Debug.Print "＊＊＊ファイル情報＊＊＊"
	Debug.Print "【ファイル名】" & objFile.Name
	Debug.Print "【ファイルサイズ】" & objFile.Size
	Debug.Print "・"
	Debug.Print "・"
	Debug.Print "・"
	Debug.Print ""
	
	'=====================================================
	' フォルダ情報
	'=====================================================
	Dim objFolder As Object
	Set objFolder = objFSO.GetFolder(sDirPath)
	Debug.Print "＊＊＊フォルダ情報＊＊＊"
	Debug.Print "【フォルダ名】" & objFolder.Name
	Debug.Print "【フォルダサイズ】" & objFolder.Size
	Debug.Print "・"
	Debug.Print "・"
	Debug.Print "・"
	Debug.Print ""
	
	'=====================================================
	' トラック情報
	'=====================================================
	Set objFolder = CreateObject("Shell.Application").Namespace(sDirPath & "\")
	
	'特定ファイルを対象とする場合
	Set objFile = objFolder.ParseName(sFileName)    'ファイル名取り出し
	Debug.Print "【ファイルサイズ】" & objFolder.GetDetailsOf(objFile, 1)   '⇒ ファイルサイズ：4.08 MB（ファイルサイズ）
	Debug.Print "・"
	Debug.Print "・"
	Debug.Print "・"
	Debug.Print ""
	
	'フォルダ内すべてのファイルを対象とする場合
	For Each objFile In objFolder.Items
		Debug.Print "【ファイルサイズ】" & objFolder.GetDetailsOf(objFile, 1)   '⇒ ファイルサイズ：4.08 MB（ファイルサイズ）
		Debug.Print "・"
		Debug.Print "・"
		Debug.Print "・"
	Next
	
	Set objFSO = Nothing
	Set objFolder = Nothing
	Set objFile = Nothing
End Sub
```

# トラック情報（GetDetailsOf）
## プロパティ（プロパティはＯＳのバージョンによって異なる。以下のコードで取得する。）

| 第二引数 | 説明                   |>         | データ型 | 出力例                          |
|:---|:---|:---|:---|:---|
| 1        | ファイルサイズ         | vbString | 文字列型 | 4.08 MB                         |
| 2        | ファイルの種類         | vbString | 文字列型 | MPEG layer 3                    |
| 3        | 更新日時               | vbString | 文字列型 | 2016/10/14 6:00                 |
| ・       | ・                     | ・       | ・       | ・                              |
| ・       | ・                     | ・       | ・       | ・                              |
| ・       | ・                     | ・       | ・       | ・                              |

## プロパティ情報取得コード
```vba
'GetDetailsOf()の詳細情報（要素番号、タイトル情報、型名、データ）を取得する
Public Sub GetDetailsOfGetDetailsOf()
	Dim sTrgtFolderPath As String
	Dim sTrgtFileName As String
	Dim sLogFilePath As String
	sTrgtFolderPath = "Z:\300_Musics\200_Reggae@Jamaica\Artist\Alaine\Sacrifice"
	sTrgtFileName = "03 Ride Featuring Tony Matterhorn.MP3"
	sLogFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\track_title_names.txt"
	
	Dim objFolder As Object
	Set objFolder = CreateObject("Shell.Application").Namespace(sTrgtFolderPath & "\")
	Dim objFile As Object
	Set objFile = objFolder.ParseName(sTrgtFileName)
	
	Open sLogFilePath For Output As #1
	Print #1, "[Idx] " & Chr(9) & "[TypeName]" & Chr(9) & "[Title]" & Chr(9) & "[Data]"
	Dim i As Long
	For i = 0 To 400
		Print #1, _
			i & Chr(9) & _
			TypeName(objFolder.GetDetailsOf(objFile, i)) & Chr(9) & _
			objFolder.GetDetailsOf("", i) & Chr(9) & _
			objFolder.GetDetailsOf(objFile, i)
	Next i
	Close #1
End Sub
```
