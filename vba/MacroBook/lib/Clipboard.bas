Attribute VB_Name = "Clipboard"
Option Explicit

'Win32API宣言
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hData As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlag As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'本来はＣ言語用の文字列コピーだが、２つ目の引数をStringとしているので変換が行われた上でコピーされる。
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long

'定数宣言
Public Const GMEM_MOVEABLE         As Long = &H2
Public Const GMEM_ZEROINIT         As Long = &H40
Public Const GHND                  As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Public Const CF_TEXT               As Long = 1
Public Const CF_OEMTEXT            As Long = 7

' ==================================================================
' = 概要    クリップボードにテキストをコピー（Win32Apiを使用）
' = 引数    sText       String  [in]  コピー対象文字列
' = 戻値                Boolean       コピー結果
' = 覚書    Win32APIを使用する。
' =         ※ クリップボードは DataObject の PutInClipboard でも利用
' =            可能｡しかし､DataObject は参照設定が必要なうえ､特定のク
' =            リップボード形式には貼り付けされない｡（CF_UNICODETEXT
' =            のみで CF_TEXTへは貼り付けされない）
' =            上記のように DataObject を使用したくない場合に本関数
' =            を利用すること｡
' ==================================================================
Public Function CopyText( _
    sText As String _
) As Boolean
    Dim hGlobal As Long
    Dim lTextLen As Long
    Dim p As Long
    
    '戻り値をとりあえず、Falseに設定しておく。
    If OpenClipboard(0) <> 0 Then
        If EmptyClipboard() <> 0 Then
            lTextLen = LenB(sText) + 1 '長さの算出(本来はUnicodeから変換後の長さを使うほうがよい)
            hGlobal = GlobalAlloc(GHND, lTextLen) 'コピー先の領域確保
            p = GlobalLock(hGlobal)
            Call lstrcpy(p, sText) '文字列をコピー
            Call GlobalUnlock(hGlobal) 'クリップボードに渡すときにはUnlockしておく必要がある
            Call SetClipboardData(CF_TEXT, hGlobal) 'クリップボードへ貼り付ける
            Call CloseClipboard 'クリップボードをクローズ
            CopyText = True 'コピー成功
        Else
            CopyText = False
        End If
    Else
        CopyText = False
    End If
End Function
