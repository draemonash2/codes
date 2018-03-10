# アドイン
- [VbePlus](http://www.vector.co.jp/soft/dl/win95/prog/se176543.html)
	- インストール後、VBEのツールバー→アドイン→アドインマネージャ
	- VbePlus のロード方法を「起動時/ロード」に変更
- [MZTools](http://zabi.0am.jp/?p=2676)
	- インストールしても表示されない場合、HKEY\_CURRENT\_USER に必要なレジストリキーが登録されていない可能性がある。
	- Administrator の HKEY\_CURRENT\_USER からコピーすること。

# Tips
- モジュール名と関数名を同じにすることはできない。
- 二次元配列の再定義について
	- 一次元目の要素数は変更できない。二次元目のサイズのみ拡張できる。
	- 二次元目の要素数を変更した場合、一次元目にまで影響する。&br()（ex. (0, 1)、(1, 1) の要素数を持つ配列に対して (1, 1) ⇒ (1, 3) と要素数を変更した場合、(0, 3)、(1, 3) の配列となる）
- エラー「定数式が必要です」について
	- デバッグを途中で停止すると、設定しているはずの値が未定義扱いとなりコンパイルエラーが発生することがある。
	- 対策は ENUM 定義名を編集して再度戻す。
- Sheets と Worksheets の違い
	- Excel のシートには複数の種類がある。（ex.グラフシート、モジュールシート、ワークシート…）
	- Sheets は「上記のすべてを包含した」オブジェクト
	- Worksheets は「ワークシートのみ」のオブジェクト
- IntegerとLongどちらが良い？
	- 型の違いによる速度比較
		- ５千万回値の代入を行った結果、Integer型：515ms、Long型：465ms かかった。（10%程Long型の方が早い）
			- 環境：Windows 7 64bit
			- オフィス：Office 2013 32bit
		- 理由は動作環境が 32bit Office のため、 4Byte 型変数の方がアクセスが早いためと思われる。
	- 型の違いによるサイズ比較
		- Integer型変数（2Byte）なら１GB満たすまで５億個変数宣言できる
		- Long型変数（4Byte）なら１GB満たすまで2.5億個変数宣言できる（容量の圧迫は気にならない）
	- 上記を踏まえると Long 型を使用するのが良い！（アクセスも早いし、容量も気にするほど大きくないため）
- Const配列の作り方
	- Const配列は VBA では定義できないが、Const 文字列を Split することで実現可能。（書き換えられてしまうが。。）
	- Const C\_AAA = "Hello!,World!" : asArray = Split(C\_AAA, ",")
- CreateObject 関数について
	- 作成したオブジェクトは、使用後に必ず Nothing を設定すること！
	- オブジェクトが解放されなくなりメモリに残り続けることになる！ 
- 連続した空白を持つ文字列を Split 関数で分割した場合、空文字列の要素ができてしまう問題
	- 事前に連続した空白を一つにまとめてから Split する
		```vba
		'連続した空白を一つにまとめる
		Do While InStr(sTrgtLine, "  ")
			sTrgtLine = Replace$(sTrgtLine, "  ", " ")
		Loop
		sTrgtLine = Trim$(sTrgtLine)
		asSplitLine = Split(sTrgtLine, " ")
		```
- クラスモジュールのメリット・デメリット
	- デメリット
		- インスタンス化したオブジェクトを他クラスで参照できない。
			- ⇒オブジェクトを生成したクラスをグローバルにすれば良い
		- Property を使用する場合、コード量が多い
		- 構造体が使えない
			- クラスは入れ子にすることができる。
				```vba
				'Classモジュール ClassString
				Public Function GetDup( _
					ByVal sVal As String
				) As String
					GetDup = sVal & sVal
				End Function
				 
				'Classモジュール ClassCom
				Public STR As New ClassString
				 
				'標準モジュール Module1
				Public C As New ClassCom
				Sub Test02()
					Debug.Print C.STR.GetDup("はい")
				End Sub
				```
		- 列挙体が使えない
	- メリット
		- グローバル変数を隠蔽できる。
			- グローバル変数を変更する場合の影響範囲をクラス内に制限できる
		- 別クラスなら同じ関数名を使える。
			- インターフェース
				```vba
				Private Sub Class_Initialize()
				End Sub
				 
				Public Function Execute() As String
				End Function
				```
			- クラス１
				```vba
				Implements ClassCmn
				 
				Dim a As String
				 
				Private Sub Class_Initialize()
					a = "A"
				End Sub
				 
				Public Function ClassCmn_Execute() As String
					ClassCmn_Execute = a
				End Function
				```
			- クラス２
				```vba
				Implements ClassCmn
				 
				Dim b As String
				 
				Private Sub Class_Initialize()
					b = "B"
				End Sub
				 
				Public Function ClassCmn_Execute() As String
					ClassCmn_Execute = b
				End Function
				```
			- 使用箇所
				```vba
				Sub test()
					Dim clChkExecTable(1) As ClassCmn
				 
					Set clChkExecTable(0) = New Class1
					Set clChkExecTable(1) = New Class2
				 
					Debug.Print clChkExecTable(0).Execute
					Debug.Print clChkExecTable(1).Execute
				 
				End Sub
				```
		- 初期化処理呼び出し不要。
- 高速化方法まとめ
	- csv ファイル取り込み方法
	- 一覧へのアクセスは Dictionary を使う。
	- セルへのアクセスは variant 型変数へ代入してから書き戻す。
	- グラフ生成は極力オプションを減らす。
	- 正規表現検索は極力使わない。InStr で代用（先頭にあるかどうかは InStr() = 0 で判定する）
	- 空白埋めは一つ一つセルに代入せず、空白セル選択からの一括値代入。
	- With を使う。
	- Integer ではなく Long を使う。
- 配列・Collection・Dictionary の速度
	- 詳細な測定結果は添付ファイル「Array・Collection・Dictionary速度比較.xlsm」参照
		- 配列
			- 全体的に爆速！インデックスアクセスだけなら配列を使うこと！
		- Dictionary
			- インデックスアクセスは遅すぎるので、絶対インデックスアクセスするな！必ず For Each でアクセスすること！
		- Collection
			- 大きな要素番号にアクセスする際は遅くなる！ので、要素数が予想できない場合は基本使わない。ただし、Key アクセスの場合は問題なし！
- CPU 使用率抑える方法
	- 長い処理の間に sleep 1 を追加する。
	- 処理は長くなるが、バックグラウンドでの処理がはかどる。
- ファイルオープン「バイナリモード」について
	- 改行コードなどを区切り文字として扱わず、１バイト単位でファイルの先頭から逐次読み書きする。
	- バイナリファイルの読み書きをする際に使用する。
- Like 演算子について
	- 正規表現の「ように」比較できる。
	- 使用例：If sAddress Like "[!東京,横浜,千葉]\*" Then ～
		- 「東京、横浜、千葉ではない住所」の場合～
			
| 記号         | 意味                                        | 使用例      | マッチする文字列          |
|:---|:---|:---|:---|
| ?            | 任意の1文字                                 | たな?       | たなか、たなべ、たなし(など)       |
| \*           | 0個以上の任意の文字                         | たか\*      | たかだ、たかなか、たかやなぎ(など) |
| #            | 1文字の数値(0～9)                           | ##          | 01、26、95(など) |
| [charlist]   | charlistに指定した文字の中の1文字           | [A-F]       | A、B、C、D、E、F |
| [!charlist]  | charlistに指定した文字の中に含まれない1文字 | [!A-F]      | G 、H、I(など) |

- On Error の入れ子について
	- On Error は関数を跨いだ入れ子をしても問題なし。
	- ただし、関数内の入れ子は適用されない。
		```vb
		Sub main()
			On Error Resume Next
			On Error Resume Next
			Call subfunction
			Debug.Print 5 / 0 '⇒実行時エラー無し
			On Error GoTo 0
		   'Debug.Print 5 / 0 '⇒実行時エラーが発生してプログラムが停止する。
			On Error GoTo 0
		End Sub
		Sub subfunction()
			On Error Resume Next
			Debug.Print 5 / 0 '⇒実行時エラー無し
			On Error GoTo 0
			Debug.Print 5 / 0 '⇒実行時エラー無し
		End Sub
		```
- 大きさの単位「ポイント」と「ピクセル」について
	- ポイント：Microsoftが定義する単位（おそらく）。8.38 ポイントはMSゴシック11ポイント(半角) で8文字と少し表示できる幅。
	- ピクセル：ディスプレイの表示やプリンタの出力を構成する最小単位（小さな点）

# VBE 設定
- Excel リボンに「開発」を追加
	- ツール ⇒ オプション ⇒ リボンユーザー設定にて [\*] 開発
- 変数宣言を強制 (変数名が誤りでも実行時までエラーが発生しないため変更)
	- ツール ⇒ オプション ⇒ 編集 ⇒ [\*] 変数の宣言を強制する
- 自動構文チェックを無視 (改行の度に警告ウィンドウが発生するため)
	- ツール ⇒ オプション ⇒ 編集 ⇒ [\_] 自動構文チェック
- エディタ文字・背景色の変更
	- ツール ⇒ オプション ⇒ エディタ ⇒ コードの表示色 を良しなに…
- 複数行コメントアウトボタン設置
	- 表示 ⇒ ツールバー ⇒ [\*] 編集 ⇒ オプション

# VBE ショートカットキー

| 項目 | キー配置 |
|:---|:---|
| IntelliSense 実行 | Ctrl+Space |
| IntelliSense 選択 | Tab        |
| マクロ強制終了    | Ctrl+Break |
| 定義へ移動（≒タグジャンプ）  | Shift+F2 |
| 元の場所へ移動（≒タグジャンプ）  | Ctrl+Shift+F2 |

# 構文
## 「～」は改行を示す。
- 【変数強制定義】Option Explicit
- 【変数/配列定義】Dim aVal(5) As Integer '要素数は０オリジン。左の例では要素数６の配列が作成される
- 【定数定義】Const NUM As Integer = 1
- 【構造体定義】Type T\_XXX ～ iVal1 As Integer ～ iVal2 As Integer ～ End Type
- 【関数定義】Private Function FuncA ( sVal1 As String, sVal2 As Integer ) ～ End Function
- 【関数呼出】Call Func()
- 【関数エラー値返却】Dim vRetVal As Variant ～ vRetVal = CVErr(xlErrRef) '#VALUEを返却
- 【マクロ定義】Public Sub SubA ( sVal1 As String, sVal2 As Integer ) ～ End Sub
- 【列挙型定義】Enum E\_XXX ～ NUM1 ～ NUM2 ～ End Enum
- 【ブロック脱出（Sub/Function/For/Do）】Exit (Sub|Function|For|Do)
- 【連続コマンド実行】Dim sStr As String : sStr = "abc"
- 【一時停止】Stop
- 【クラスインスタンス生成】Dim cPrfrmMes As New PerformanceMeasurement
- 【クラスインスタンス破棄】Set cPrfrmMes = Nothing
- 【プログラム終了】End
- 【関数エラー値返却】Dim vRetVal As Variant ～ vRetVal = CVErr(xlErrRef) '#VALUEを返却（エラー値の詳細は [こちら](https://msdn.microsoft.com/ja-jp/library/office/ff839168.aspx) 参照）

- 【if】If iVal = 1 Or iVal = 2 Then ～ ElseIf iVal = 3 Then ～ Else ～ End If
- 【if（空オブジェクト確認）】If objTest Is Nothing Then ～ Else ～ End If
- 【switch】Select Case iVal ～ Case 1 ～ Case Else ～ End Select
- 【for】For iVal1 = 1 To 3 [Step 1] ～ Next Val
- 【for each】For Each Value in Values ～処理～ Next
- 【while】Do ～(条件式＝真)～ Loop While 条件式
- 【do while】Do While 条件式 ～(条件式＝真)～ Loop
- 【do until】Do Until 条件式 ～(条件式＝偽)～ Loop
- 【with】With オブジェクト名 ～ End With
- 【コメント】'コメント
- 【入力】sStr = InputBox( "テキストを入力してください", "title", "default value" )
- 【出力１】MsgBox "Hello world", vbExclamation '第二引数入力時に候補が表示される
- 【出力２】Debug.Print "Hello world"
- 【チェック処理】Debug.Assert 条件式 'Falseの場合、処理を一時停止
- 【確認処理】Dim vAnswer As Variant ～ vAnswer = MsgBox("処理を継続しますか？", vbOKCancel, "タイトル") '第二引数は表示ボタンの種類を指定。ボタンの種類は [こちら](http://www.kanaya440.com/contents/script/vbs/function/others/msgbox.html) 参照。

- 【正規表現】サンプルコード参照

- 【コレクション 定義】Dim cTrgtPaths As Variant ～ Set cTrgtPaths = CreateObject("System.Collections.ArrayList")
	- 【コレクション 追加】cTrgtPaths.Add "c:\test\a.txt"
	- 【コレクション 値取り出し（単一）】cTrgtPaths.Item(0) 'c:\test\a.txt（０オリジン）
	- 【コレクション 値取り出し（ループ）】Dim vTrgtPath As Variant ～ For Each vTrgtPath In cTrgtPaths ～ MsgBox vTrgtPath ～ Next
	- 【コレクション 要素数取得】cTrgtPaths.Count '要素数（末尾の要素番号ではない）
	- 【コレクション 削除】cTrgtPaths.Remove "c:\test\b.xlsx" '要素の値を指定する。要素番号では削除できない。
	- 【コレクション 挿入】cTrgtPaths.Insert 2, "c:\test\e.ppt" '要素番号2へ挿入される（元要素番号2以降が一要素ずれる）
	- 【コレクション ソート】cTrgtPaths.Sort
	- 【コレクション 配列変換】Dim avTrgtPaths As Variant ～ avTrgtPaths = cTrgtPaths.ToArray() 'Variant型配列に変換
	- 【コレクション 全要素削除】cTrgtPaths.Clear

- 【連想配列 定義】Dim oPriceOfFruit As Object ～ Set oPriceOfFruit = CreateObject("Scripting.Dictionary")&br()（CreateObject 関数でインスタンス化した場合、VBA エディターでの自動補完（インテリセンス）が働かない！）
	- 【連想配列 キー/項目追加】oPriceOfFruit.Add("リンゴ", "100円")
	- 【連想配列 存在確認】oPriceOfFruit.Exists("リンゴ")
	- 【連想配列 キー取得（For Each）】For Each vKey In oPriceOfFruit ～ Debug.print vKey ～ Next 'vKey は variant 型
	- 【連想配列 項目取得（キー）】oPriceOfFruit.Item("リンゴ")
	- 【連想配列 キー取得（インデックス）】oPriceOfFruit.Keys()(0) '０オリジン（アクセスが遅すぎるので注意！）
	- 【連想配列 項目取得（インデックス）】oPriceOfFruit.Items()(0) '０オリジン（アクセスが遅すぎるので注意！）
	- 【連想配列 キー置換】oPriceOfFruit.Key("リンゴ") = "りんご"
	- 【連想配列 キー関連付け】oPriceOfFruit.Item("リンゴ") = "200円"
	- 【連想配列 キー/項目数取得】oPriceOfFruit.Count
	- 【連想配列 キー/項目削除】oPriceOfFruit.Remove("リンゴ") '指定されたキーが存在しない場合はエラー
	- 【連想配列 キー/項目全削除】oPriceOfFruit.RemoveAll
	- 【連想配列 配列変換（項目）】asFruitPrice = oPriceOfFruit.Items 'Variant型配列、０オリジン
	- 【連想配列 配列変換（キー）】asFruitName = oPriceOfFruit.Keys 'Variant型配列、０オリジン
	- 【連想配列 設定変更】oPriceOfFruit.CompareMode = vbBinaryCompare '大/小文字区別 /vbTextCompare（大/小文字区別しない

- 【エラー設定】On Error Resume Next
- 【エラー解除】On Error Goto 0
- 【エラー番号】Err.Number
- 【エラー内容】Err.Description
- 【エラーラベル】On Error GoTo ErrorLabel 'ErrorLabelはラベル名
- 【ラベル定義】ErrorLabel:

- 【WScriptShellObject 取得】Dim objWshShell ～ Set objWshShell = CreateObject("WScript.Shell")
	- 【バッチ実行①】objWshShell.Exec("C:\test.bat") 'Execは標準入出力できるが、WSH 5.6以降からしかサポートされていないので注意
	- 【バッチ実行②】objWshShell.Run "C:\test.bat", 0, True '第二引数：ウィンドウの表示スタイル（ウィンドウを非表示、別のウィンドウをアクティブ）、第三引数：プログラムの実行が終了するまでスクリプトを待機させるかどうか（詳細は[こちら](https://msdn.microsoft.com/ja-jp/library/cc364421.aspx)）
	- 【レジストリ読込】objWshShell.RegRead("HKCU\WshTest\Test1")
	- 【レジストリ書込】objWshShell.RegWrite("HKCU\WshTest\Test1", "test", "REG\_SZ") 'キー/値,設定値,データ型
	- 【環境変数 値取得】objWshShell.ExpandEnvironmentStrings( "%MYPATH\_CODES%" )
	- 【特殊フォルダのパス取得】objWshShell.SpecialFolders("Desktop") 'デスクトップフォルダ
		- 取得できるフォルダは「AllUsersDesktop」 「AllUsersStartMenu」 「AllUsersPrograms」 「AllUsersStartup」 「Desktop」 「Favorites」 「Fonts」 「MyDocuments」 「NetHood」 「PrintHood」 「Programs」 「Recent」 「SendTo」 「StartMenu」 「Startup」 「Templates」
	- 【ショートカット作成】With objWshShell.CreateShortcut( "c:\test\src.txt.lnk" ) ～ .TargetPath = "c:\test\dst.txt" ～ .Save ～ End With
	- 【ショートカット 指示先パス取得】objWshShell.CreateShortcut( "c:\test\src.txt.lnk" ).TargetPath '参照だけでなく変更も可
	- 【ショートカット 指示先パス更新】With objWshShell.CreateShortcut( "c:\test\src.txt.lnk" ) ～ .TargetPath = "c:\test\dst2.txt" ～ .Save ～ End With

- 【FileSystemObject 取得】Dim objFSO As Object ～ Set objFSO = CreateObject("Scripting.FileSystemObject")
	- 【ファイルコピー（自ブック）】objFSO.CopyFile ThisWorkbook.FullName, "c:\temp\test.xlsm"
	- 【ファイル コピー①】objFSO.CopyFile "C:\codes\a.txt", "C:\codes\test\" '<src> <dst> [<overwrite>] 、<dst>の末尾に "\" をつけること！
	- 【ファイル コピー②】objFSO.CopyFile "C:\codes\a.txt", "C:\codes\test\a.txt" '<src> <dst> [<overwrite>]
	- 【ファイル 削除】objFSO.DeleteFile "c:\test", True
	- 【ファイル 移動/リネーム】objFSO.MoveFile "C:\codes\src.txt", "C:\codes\dst.txt"
	- 【ファイル 存在確認①】If Dir("C:\Book1.xlsx") <> "" Then ～(存在)～ Else ～(非存在)～ End If
	- 【ファイル 存在確認②】objFSO.FileExists("c:\codes\a.txt") 'True
	- 【ファイル 情報取得】objFSO.GetFile( "C:\codes\a.txt" ).Attributes '32 (※)値の意味は [[【ファイル・フォルダ情報取得】]] 参照
	- 【ファイル 隠しファイル化】objFSO.GetFile( "C:\codes\a.txt" ).Attributes = 2
	- 【ファイル 絶対パス取得】objFSO.GetAbsolutePathName( "C:\codes\a.txt" ) ' C:\codes\a.txt
	- 【ファイル ドライブ名取得】objFSO.GetDriveName( "C:\codes\a.txt" ) ' C:
	- 【ファイル ファイル名取得】objFSO.GetFileName( "C:\codes\a.txt" ) ' a.txt
	- 【ファイル ファイルベース名取得】objFSO.GetBaseName( "C:\codes\a.txt" ) ' a
	- 【ファイル 拡張子取得】objFSO.GetExtensionName( "C:\codes\a.txt" ) ' txt
	- 【ファイル 親フォルダパス取得】objFSO.GetParentFolderName( "C:\codes\a.txt" ) ' C:\codes
	- 【フォルダ コピー】objFSO.CopyFolder "C:\codes\src", "C:\codes\dst", True '配下フォルダ/ファイルも丸ごとコピー
	- 【フォルダ 削除】objFSO.DeleteFolder "C:\codes\test", True '配下フォルダ/ファイルも丸ごと削除
	- 【フォルダ 作成】objFSO.CreateFolder( "C:\codes\test" ) '親フォルダがない場合、エラーになる
	- 【フォルダ 移動/リネーム】objFSO.MoveFolder "C:\codes\src", "C:\codes\dst" '配下フォルダ/ファイルも丸ごと移動/リネーム
	- 【フォルダ 情報取得】objFSO.GetFolder( "C:\codes" ).Attributes '32 (※)値の意味は [[【ファイル・フォルダ情報取得】]] 参照
	- 【フォルダ 存在確認】objFSO.FolderExists( "C:\codes" ) 'True
	- 【フォルダ 親フォルダパス取得】objFSO.GetParentFolderName( "C:\codes\src" ) ' C:\codes

- 【ＴＸＴファイルオープン/クローズ】Open ファイル名 For [Input|Output|Append] As #1 ～ Close #1
- 【ＴＸＴファイル読込（一行ずつ）】Do Until EOF(1) ～ Line Input #1, 文字列変数 ～ Loop
- 【ＴＸＴファイル読込（一括）】sTestFile = objFSO.GetFile(ファイルパス).OpenAsTextStream.ReadAll&br()'返却値は「配列型」でないことに注意！改行文字を含んだ「文字列型」で返却される！
- 【ＴＸＴファイル書込】Print #1, 文字列変数
- 【ＸＬＳファイルオープン/クローズ】Set wTargetBook = Workbooks.Open(sTargetBookName) ～ wTargetBook.Close SaveChanges:=True

- 【置換】Replace(文字列変数, "  ", "")
- 【文字列検索（前方）】InStr("abcabc", "bc") ' 先頭からの位置、１オリジン（いない場合は0が返る
- 【文字列検索（後方）】InStrRev("abcabc", "bc")  '末尾からの位置、１オリジン（いない場合は0が返る
- 【文字列 長さ（文字数）】Len("リンゴ") '3
- 【文字列 長さ（バイト数）】LenB("リンゴ") '6
- 【文字列結合】"abcdef" & "gh"
- 【文字列抽出 左】Left$("abcd", 3) 'abc
- 【文字列抽出 中】Mid$("abcdefgh", 3, 2) 'cd（１オリジン）
- 【文字列抽出 右】Right$("abcd", 2) 'cd
- 【文字列⇒ASCII 変換】Asc(文字)
- 【文字列の数値判定】IsNumeric( sStr ) '1⇒True、"a"⇒False、""⇒False
- 【ASCII⇒文字列 変換】Chr(ASCIIコード) (ex. Chr(Asc("①") + 1) ⇒ ② )
- 【文字列繰り返し】"a" & String("b", 4) 'abbbb
- 【大文字化】UCase("aaa")
- 【小文字化】LCase("AAA")
- 【配列再定義】ReDim Preserve 配列名(5) '要素数は０オリジン。左の例では要素数６の配列が作成される
- 【配列最大要素数】UBound(配列名) '返却値は０オリジン。返却値が３の場合、0～3 の配列であることを示す
- 【要素数０（未初期化）/要素数１配列判定】If Sgn(asStr) = 0 Then ～未初期化配列～ Else ～要素数１配列～ End If
- 【配列 結合】Join(配列, ",")
- 【配列 分割】文字列配列 = Split("aaa,bbb,ccc", ",") 
- 【型取得（文字列）】TypeName("Test") 'String
- 【型取得（値）】VarType("Test") '8（値の詳細は [こちら](http://www.kanaya440.com/contents/script/vbs/function/data/var_type.html) 参照）
- 【10⇒16進数変換】文字列変数 = Hex(734)
- 【16⇒10進数変換１】Long変数 = CLng("&H" & "FA")
- 【16⇒10進数変換２】Int変数 = CInt("&H" & "FA")
- 【符号つき16進数表現】&HFFF0
- 【符号なし16進数表現】&HFFF0&
- 【文字列⇒数値変換】Val(文字列式)
- 【数値⇒文字列変換】Str(数値)
- 【改行】( vbNewLine | vbCr | vbLf | vbCrLf )
- 【少数 正数 切り捨て①】Fix( 99.224 ) '99
- 【少数 正数 切り捨て②】Int( 99.224 ) '99
- 【少数 負数 切り捨て①】Fix( -99.224 ) '-99 (※)負数の場合は切り上げ
- 【少数 負数 切り捨て②】Int( -99.224 ) '-100 (※)負数の場合は切り下げ
- 【少数 正数 四捨五入（第一位）】Round( 99.555, 0 ) '100
- 【少数 正数 四捨五入（第二位）】Round( 99.555, 1 ) '99.6
- 【少数 正数 四捨五入（第三位）】Round( 99.555, 2 ) '99.56
- 【少数 負数 四捨五入（第一位）】Round( -99.555, 0 ) '-100
- 【少数 負数 四捨五入（第二位）】Round( -99.555, 1 ) '-99.6
- 【少数 負数 四捨五入（第三位）】Round( -99.555, 2 ) '-99.56
- 【少数 正数 切り上げ（第一位）】Round( 99.224 + 0.5, 0 ) '100
- 【少数 正数 切り上げ（第二位）】Round( 99.224 + 0.05, 1 ) '99.3
- 【少数 負数 切り上げ（第一位）】Round( -99.224 - 0.5, 0 ) '-100
- 【少数 負数 切り上げ（第二位）】Round( -99.224 - 0.05, 1 ) '-99.3
- 【少数 正数 切り下げ（第一位）】Round( 99.224 - 0.5, 0 ) '99 （＝切り捨て）
- 【少数 正数 切り下げ（第二位）】Round( 99.224 - 0.05, 1 ) '99.2
- 【少数 負数 切り下げ（第一位）】Round( -99.224 + 0.5, 0 ) '-99 （＝切り捨て）
- 【少数 負数 切り下げ（第二位）】Round( -99.224 + 0.05, 1 ) '-99.2

- 【現在時刻取得】Now 'YYYY/DD/MM HH:MM:SS
- 【現在年月日取得】Date 'YYYY/MM/DD
- 【0:00から現在までの経過時間（秒数）】Timer '49229.781（13:40:29 .781）
- 【日付比較】If DateDiff("s", sCmpBaseTime, sModDate ) > 0 Then ～（ sModDate が新しい）～ ElseIf DateDiff("s", sCmpBaseTime, sModDate ) < 0 Then ～（ sModDate が古い）～ Else ～（ sModDate = sCmpBaseTime）～ End If '日付は "YYYY/MM/DD" の文字列として指定する
- 【Wait処理】    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ～ Sleep 1000 'ms 単位

- 【ブック作成】Dim wTrgtBook As Workbook ～ Application.SheetsInNewWorkbook = 1 ～ Set wTrgtBook = Workbooks.Add 'SheetsInNewWorkbook は作成時のシート数を指定
- 【ブック既に開いているか確認】If wCsvBook.Name <> Dir("C:\Book1.xlsx") Then ～(処理)～ Else ～(エラー)～ End If
- 【ブック追加】Dim bAddBook As Workbook ～ Set bAddBook = Workbooks.Add
- 【ブック作成＆シートコピー】ThisWorkbook.Sheets(シート名).Copy: Set wTrgtBook = ActiveWorkbook
- 【ブック新規保存】wTrgtBook.SaveAs Filename:=sFilePath
- 【シート数取得】.Sheets.Count
- 【シート追加】Dim shAddSht As Worksheet ～ Set shAddSht = ThisWorkbook.Sheets.Add
- 【シート移動（末尾）】ThisWorkbook.Sheets(シート名).Move After:=ThisWorkbook.Sheets( ThisWorkbook.Sheets.Count )
- 【シート削除】Application.DisplayAlerts = False ～ .Sheets(シート名).Delete ～ Application.DisplayAlerts = True
- 【シート表示/非表示】.Sheets(シート名).Visible = (True|False)
- 【シート並べ替え】.Sheets(シート名).Move Before:=Sheets(1)
- 【画面表示 ON】Application.ScreenUpdating = True
- 【画面表示 OFF】Application.ScreenUpdating = False
- 【計算方法切替 自動】Application.Calculation = xlCalculationAutomatic
- 【計算方法切替 手動】Application.Calculation = xlCalculationManual
- 【ブック再計算】Application.Calculate
- 【ブック強制再計算】Application.CalculateFull
- 【確認メッセージ抑制/表示】Application.DisplayAlerts = (False|True)
- 【セル検索（行番号取得）】.Cells.Find("りんご", LookAt:=xlWhole).Row
- 【セル検索（行番号取得）】.Cells.Find("りんご", LookAt:=xlWhole).Column
- 【セル参照方法】.Cells(X, Y).Value 'I,X,Yは１オリジン
- 【セル位置取得（Ｙ軸）】.Cells(1, 1).Top '単位：ポイント
- 【セル位置取得（Ｘ軸）】.Cells(1, 1).Left '単位：ポイント&br()（セルの高さ/幅は取得できないので、次セルの Top/Left との差分で判断する必要がある）
- 【行削除】.Cells(1, 1).EntireRow.Delete Shift:=xlShiftUp
- 【行追加】Application.CutCopyMode = False ～ .Range("2:4").Insert
- 【取消線取得】.Cells(行, 列).Font.Strikethrough
- 【枠線非表示】wTrgtBook.Sheets(シート名).Activate ～ ActiveWindow.DisplayGridlines = False
- 【フォント名取得】文字列変数 = .Range("A1").Font.Name
- 【可視/不可視チェック（行）】.Range("A1").EntireRow.Hidden
- 【可視/不可視チェック（列）】.Range("A1").EntireColumn.Hidden
- 【フォントサイズ変更】.Range("A1").Font.Size = 14
- 【フォントカラー変更】.Range("A1").Font.Color = RGB(0, 255, 0)
- 【フォント太字変更】.Range("A1").Font.Bold = True
- 【フォント下線変更】.Range("A1").Font.Underline = True
- 【背景色変更】.Range("A1").Interior.Color = RGB(255, 255, 0)
- 【罫線（格子）設定】.Range("A1:C3").Borders.LineStyle = xlContinuous
- 【セル結合】.Range("A1:C3").MergeCells = True
- 【セル選択範囲内中央】.Range("A1:C3").HorizontalAlignment = ( xlCenterAcrossSelection | xlGeneral )
- 【最終行取得】数値変数 = .Cells(.Rows.Count, 列).End(xlUp).Row
- 【最終列取得】数値変数 = .Cells(行, .Columns.Count).End(xlToLeft).Column
- 【最終行取得（全列の中で最大）】数値変数 = .Sheets(シート名).UsedRange.Rows.Count + 1
- 【最終列取得（全行の中で最大）】数値変数 = .Sheets(シート名).UsedRange.Columns.Count + 1
- 【選択範囲位置取得（先頭行）】Selection(1).Row
- 【選択範囲位置取得（末尾行）】Selection(Selection.Count).Row
- 【選択範囲位置取得（先頭列）】Selection(1).Column
- 【選択範囲位置取得（末尾列）】Selection(Selection.Count).Column
- 【Rangeオブジェクトの行/列番号指定】.Range(.Cells(1, 1), .Cells(6, 3))
- 【セルコピー（書式保持）】.Range("A1:B9").Copy Destination:=ThisWorkBook.Sheets(シート名２).Range("B1")
- 【セルソート】.Range(.Cells(1, 1), .Cells(.Rows.Count, 2)).Sort Key1:=.Cells(1, 2) ,order1:=xlAscending
- 【範囲セル 値クリア（書式そのまま）】.Range("A1:A2").ClearContents
- 【範囲セル 書式クリア】.Range("A1:A2").ClearFormats
- 【範囲セル 書式貼り付け】.Range("A1:A2").PasteSpecial (xlPasteFormats) '数式：xlPasteFormulas、値：xlPasteValues、書式：xlPasteFormats、コメント：xlPasteComments、入力規則：xlPasteValidation
- 【空白セル選択】.Range("A1:CV100").SpecialCells(xlCellTypeBlanks).Select
- 【列幅変更】.Range("A1").ColumnWidth = 5 '行はRowHeight
- 【自動列幅調整】.Range(.Cells(4, 2), .Cells(9, 2)).Columns.AutoFit '行はRows
- 【自動列幅調整（全領域）】.Sheets(シート名).UsedRange.Columns.AutoFit '行はRows
- 【選択セルに対して「選択範囲内で中央」】Selection.HorizontalAlignment = xlCenterAcrossSelection
- 【コピー/切り取りモード解除】Application.CutCopyMode = False
- 【グループ化（行）】.Range( .Rows( lStrtRow ), .Rows( lLastRow ) ).Group '解除は Group → Ungroup
- 【グループ化（行）】.Range( .Columns( lStrtRow ), .Columns( lLastRow ) ).Group '解除は Group → Ungroup
- 【アウトライン設定変更（上下）】.Sheets(シート名).Outline.SummaryRow = ( xlBelow | xlAbove )
- 【アウトライン設定変更（左右）】.Sheets(シート名).Outline.SummaryColumn = ( xlRight | xlLeft )
- 【アウトライン設定変更（自動）】.Sheets(シート名).Outline.AutomaticStyles = ( True | False )

- 【ChartObject定義】Dim oChartObj As ChartObject ～ Set oChartObj = ThisWorkbook.Sheets(シート名).ChartObjects(1)
	- 【グラフ 追加】Set oChartObj = .Sheets(シート名).ChartObjects.Add( XPOS, YPOS, WIDTH, HEIGHT ) 'XPOS, YPOS, WIDTH, HEIGHTの単位はポイント
	- 【グラフ 削除】oChartObj.Delete
	- 【グラフ コピー】oChartObj.Chart.ChartArea.Copy
	- 【グラフ 移動（Ｙ軸）】oChartObj.Top = 10
	- 【グラフ 移動（Ｘ軸）】oChartObj.Left = 20
	- 【グラフ サイズ変更（幅）】oChartObj.Width = 200
	- 【グラフ サイズ変更（高さ）】oChartObj.Height = 300
	- 【グラフ 種別】oChartObj.Chart.ChartType = xlXYScatterLines 'xlXYScatterLines:折れ線付き散布図、xlLine:折れ線、...
	- 【グラフ データ範囲変更】oChartObj.Chart.SetSourceData Source:=Union(rXAxsRng, rDataRng) 'データ範囲指定
	- 【グラフ Ｘ軸 タイトル 有無】oChartObj.Chart.Axes(xlCategory).HasTitle = True
	- 【グラフ Ｘ軸 タイトル 変更】oChartObj.Chart.Axes(xlCategory).AxisTitle.Text = "Test Axis X"
	- 【グラフ Ｘ軸 目盛軸 有無】oChartObj.Chart.Axes(xlCategory).HasMajorGridlines = True
	- 【グラフ Ｘ軸 目盛軸 色】oChartObj.Chart.Axes(xlCategory).MajorGridlines.Border.Color = RGB(217, 217, 217)
	- 【グラフ Ｘ軸 目盛軸 太さ】oChartObj.Chart.Axes(xlCategory).MajorGridlines.Border.Weight = 2
	- 【グラフ Ｘ軸 目盛軸 スタイル】oChartObj.Chart.Axes(xlCategory).MajorGridlines.Border.LineStyle = (xlContinuous|xlDot|xlDouble|xlLineStyleNone|...)
	- 【グラフ Ｘ軸 補助目盛軸 〃】上記【グラフ Ｘ軸 目盛軸 ～】の「MajorGridlines」を「MinorGridlines」に変更
	- 【グラフ Ｘ軸 最小値 自動】oChartObj.Chart.Axes(xlCategory).MinimumScaleIsAuto = False
	- 【グラフ Ｘ軸 最大値 自動】oChartObj.Chart.Axes(xlCategory).MaximumScaleIsAuto = False
	- 【グラフ Ｘ軸 最小値 設定】oChartObj.Chart.Axes(xlCategory).MinimumScale = 0
	- 【グラフ Ｘ軸 最大値 設定】oChartObj.Chart.Axes(xlCategory).MaximumScale = 100
	- 【グラフ Ｘ軸 縦軸との交点】oChartObj.Chart.Axes(xlCategory).Crosses = (xlMinimum|xlMaximum|xlAutomatic)
	- 【グラフ Ｙ軸 〃】上記【グラフ Ｘ軸 ～】の「xlCategory」を「xlValue」に変更
	- 【グラフ タイトル 有無】oChartObj.Chart.HasTitle = True
	- 【グラフ タイトル 変更】oChartObj.Chart.ChartTitle.Text = "Test Title"
	- 【グラフ タイトル グラフに重ねる】oChartObj.Chart.ChartTitle.IncludeInLayout = False 'False:重ねる、True:重ねない
	- 【グラフ 凡例 有無】oChartObj.Chart.HasLegend = True
	- 【グラフ 凡例 位置】oChartObj.Chart.Legend.Position = (xlLegendPositionTop|xlLegendPositionBottom|xlLegendPositionLeft|xlLegendPositionRight|...)
	- 【グラフ 凡例 グラフに重ねる】oChartObj.Chart.Legend.IncludeInLayout = False 'False:重ねる、True:重ねない
- 【グラフ 画像として貼り付け】.Sheets(シート名).PasteSpecial Format:="図 (JPEG)", Link:=False, DisplayAsIcon:=False

- 【ワークシート関数】Application.WorksheetFunction.VLookup(.Range("C1"), .Range("A1:B7"), 2, False)

- 【フォーム ロード】Dim goPrgrsBar As New ProgressBar ～ goPrgrsBar.Show vbModeless
- 【フォーム アンロード】goPrgrsBar.Hide ～ Unload goPrgrsBar ～ Set goPrgrsBar = Nothing

- 【Collection 項目取得（キー）】cCollection.Item("リンゴ")
- 【Collection 項目取得（インデックス）】cCollection(0)
- 【Collection 項目取得（For Each）】For Each vItem In cCollection ～ Debug.print vItem ～ Next

- 【チェックボックス値取得（フォームコントロール）】lChk = ThisWorkbook.Sheets(シート名).CheckBoxes(1).Value 'On:1 Off:-4146
- 【ユーザフォーム表示中のキー操作】Private Sub xxx\_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) ～ End Sub
	- xxx はフォーカス中のフォーム名。フォーム xxx にフォーカスがある場合だけ、KeyUp イベントが発生する。
	- どのフォームにフォーカスがあっても動く KeyUp イベントを作りたい場合、全フォームに対して上記イベントを作る！

# ライブラリ
- [【ＸＬＳファイル存在確認～オープン～クローズ】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelFile.bas)
- [【ＴＸＴファイル存在確認～オープン～クローズ】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [【ＴＸＴファイル存在確認～オープン～クローズ（キャラクタセット指定）】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [【フォルダ作成（再帰処理）】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [【ファイル名一覧取得】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [【ファイル・ディレクトリ 判別】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [【ファイル・ディレクトリ 選択ダイアログ】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [【ファイル・フォルダ情報取得】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileInfo.bas)
- [【コマンド実行】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/SysCmd.bas)
- [【プログレスバー表示】]
- [【性能測定】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/StopWatch.cls)
- [【エラー処理】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/Error.bas)
- [【配列の Push, Pop 関数】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ArrayMng.bas)
- [【イミディエイトウィンドウクリア】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/VbaMng.bas)
- [【シート一覧】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [【ツリー図オブジェクト生成】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [【ファイル名・フォルダ名抽出関数】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/StringMng.bas)
- [【セル範囲文字結合/分割関数】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [【セル属性取得関数】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [【ビット演算関数】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [【キー送信】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/SendKeys.bas)
- [【特殊貼り付け】](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/SpecialPaste.bas)

# サンプルコード
- [[【プログラムテンプレート】]]
- 【シート存在確認】
	```vba
	'シート存在確認
	Dim bIsSheetExist As Boolean
	bIsSheetExist = False
	For Each wSheet In ThisWorkbook.Worksheets
		If wSheet.Name = sTrgtSheetName Then
			bIsSheetExist = True
		Else
			'Do Nothing
		End If
	Next wSheet
	```
- 【セル検索（存在しない場合を考慮）】
	- セル検索
		```vba
		Dim rFindResult As Range
		Dim sSrchKeyword As String
		Dim lSrchCellRow As Long
		Dim lSrchCellClm As Long
		
		With ThisWorkbook.Sheets(INPUT_SHEET_NAME)
			sSrchKeyword = KYWD_INPUT_DIR_PATH
			Set rFindResult = .Cells.Find(sSrchKeyword, LookAt:=xlWhole)
			If rFindResult Is Nothing Then
				MsgBox sSrchKeyword & "が見つかりませんでした"
			Else
				lSrchCellRow = rFindResult.Row
				lSrchCellClm = rFindResult.Column
			End If
		End With
		```
	- 直近セル情報取得
		```vba
		Public Type T_EXCEL_NEAR_CELL_DATA
			bIsCellDataExist As Boolean
			lRow As Long
			lClm As Long
			sCellValue As String
		End Type
		
		Public Function GetNearCellData( _
			ByVal shTrgtSht As Worksheet, _
			ByVal sSrchKey As String, _
			ByVal lRowOffset As Long, _
			ByVal lClmOffset As Long _
		) As T_EXCEL_NEAR_CELL_DATA
			Dim rFindResult As Range
			Set rFindResult = shTrgtSht.Cells.Find(sSrchKey, LookAt:=xlWhole)
			If rFindResult Is Nothing Then
				GetNearCellData.bIsCellDataExist = False
			Else
				GetNearCellData.bIsCellDataExist = True
				GetNearCellData.lRow = rFindResult.Row + lRowOffset
				GetNearCellData.lClm = rFindResult.Column + lClmOffset
				GetNearCellData.sCellValue = shTrgtSht.Cells( _
													GetNearCellData.lRow, _
													GetNearCellData.lClm _
												 ).Value
			End If
		End Function
		```
- 【セル重複削除】
	```vba
	For iRowIdx = .Cells(Rows.Count, 1).End(xlUp).Row To 2 Step -1
		For iChkTargetRowIdx = iRowIdx - 1 To 2 Step -1
			If .Cells(iRowIdx, 1) = .Cells(iChkTargetRowIdx, 1) Then
				.Cells(iChkTargetRowIdx, 1).EntireRow.Delete Shift:=xlShiftUp
			Else
				'Do Nothing
			End If
		Next iChkTargetRowIdx
	Next iRowIdx
	```
- [[【グラフ作成＆画像変換】]]
- 【テキストボックス作成】
	```vba
	Dim spTxtBox As Shape
	Set spTxtBox = shTrgtSht.Shapes.AddTextbox( _
						Orientation:=msoTextOrientationHorizontal, _
						Left:=10, _
						Top:=10, _
						Width:=10, _
						Height:=10 _
				   )
	With spTxtBox
		.TextFrame.Characters.Text = "エラーフレームなし"
		.TextFrame.AutoSize = True
	End With
	```
- 【正規表現】
	- 正規表現と文字列検索 InStr の速度についての考察…
		- マッチングを 500000 回繰り返すと、正規表現は約 2800 ms、InStr では約 50 ms。
		- 正規表現は 56 倍遅い！
	- 参考URL: http://officetanaka.net/excel/vba/tips/tips38.htm
		```vba
		Dim sSearchPattern As String
		Dim sTargetStr As String
		 
		Dim iMatchIdx As Integer
		Dim iSubMatchIdx As Integer
		 
		Dim oMatchResult As Object
		Dim oRegExp As Object
		Set oRegExp = CreateObject("VBScript.RegExp")
		 
		sSearchPattern = "(\w+)\((\w+) (\w+)\)"
		sTargetStr = "TestFunc01(char aaa) TestFunc02(int bbb)"
		 
		oRegExp.Pattern = sSearchPattern               '検索パターンを設定
		oRegExp.IgnoreCase = True                      '大文字と小文字を区別しない
		oRegExp.Global = True                          '文字列全体を検索
		Set oMatchResult = oRegExp.Execute(sTargetStr) 'パターンマッチ実行
		 
		Debug.Print oMatchResult.Count
		For iMatchIdx = 0 To oMatchResult.Count - 1
			Debug.Print oMatchResult(iMatchIdx).SubMatches.Count
			For iSubMatchIdx = 0 To oMatchResult(iMatchIdx).SubMatches.Count - 1
				Debug.Print oMatchResult(iMatchIdx).SubMatches(iSubMatchIdx)
			Next iSubMatchIdx
		Next iMatchIdx
		```
- 【連想配列】
	- 連想配列を使用することにより、膨大な件数(※)の検索が劇的に早くなる。
		- ※ 検索対象 15 件以上（それ未満だと配列検索のほうが早い。.Exists API のオーバーヘッドのせい？）
			```vba
			Dim oPriceOfFruit As Object
			Set oPriceOfFruit = CreateObject("Scripting.Dictionary")
			 
			'登録
			oPriceOfFruit.Add "リンゴ", "100円"
			oPriceOfFruit.Add "イチゴ", "400円"
			oPriceOfFruit.Add "メロン", "1000円"
			 
			'存在確認
			If oPriceOfFruit.Exists("リンゴ") Then
				Debug.Print "Exists!"
			Else
				Debug.Print "Not Exists!"
			End If
			 
			Set oPriceOfFruit = Nothing
			```

# その他
- 【型一覧】

| データ型 | 名称 | 消費メモリ | 格納できる範囲 |
|:---------|:-----|:-----------|:---------------|
| Integer | 整数型 | 2バイト | -32,768 ～ 32,767 |
| Long | 長整数型 | 4バイト | -2,147,483,648 ～ 2,147,483,647 |
| Single | 単精度浮動小数点数型 | 4バイト | -3.402823E38 ～ -1.401298E-45(負の値) 1.401298E-45 ～ 3.402823E38(正の値) |
| Double | 倍精度浮動小数点数型 | 8バイト | -1.79769313486232E308 ～ -4.94065645841247E-324(負の値) 4.94065645841247E-324 ～ 1.79769313486232E308(正の値) |
| Currency | 通貨型 | 8 バイト | -922,337,203,685,477.5808 ～ 922,337,203,685,477.5807 |
| String | 文字列型 | 2バイト | 最大約20億文字まで |
| Date | 日付型 | 8 バイト | 西暦100 年1月1日～西暦9999年12月31日までの日付と時刻 |
| Object | オブジェクト型 | 4 バイ | オブジェクトを参照するデータ型 |
| Variant | バリアント型 | 16バイト | 可変長の文字列型の範囲と同じ。 |
| Boolean | ブール型 | 2 バイト | 真 (True) または偽 (False) |

- 【演算子】

| 演算子 | 意味 | 変数への代入例 |
|:-------|:-----|:---------------|
| ＋ | 加算する | i = 15 + 5 ( ｉ の値は20) |
| － | 減算する | i = 15 - 5 ( ｉ の値は10) |
| ＊ | 乗算する | i = 5 * 4 ( ｉ の値は20) |
| ／ | 除算する | i = 15 / 5 ( ｉ の値は3) |
| ￥ | 除算の商 | i = 15 \ 2 ( ｉ の値は7) |
| Mod | 除算の余り | i = 15 Mod 2 ( ｉ の値は1) |
| ＾ | べき乗する | i = 2 ^ 5 ( ｉ の値は32) |

- 【エラー種別】
	- ex) CVErr(xlErrNum)
	
| 定数       | エラー番号 | セルのエラー値 |
|:-----------|:-----------|:---------------|
| xlErrDiv0  | 2007       | #DIV/0!        |
| XlErrNA    | 2042       | #N/A           |
| xlErrName  | 2029       | #NAME?         |
| XlErrNull  | 2000       | #NULL!         |
| XlErrNum   | 2036       | #NUM!          |
| XlErrRef   | 2023       | #REF!          |
| XlErrValue | 2015       | #VALUE!        |
