Attribute VB_Name = "Mng_Clipboard"
Option Explicit

' clipboard library v2.0

#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
#Else
Private Declare Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#End If
'GlobalALock
Private Const GHND = &H42
Private Const CF_TEXT = &H1
Private Const CF_LINK = &HBF00
Private Const CF_BITMAP = 2
Private Const CF_METAFILE = 3
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const MAXSIZE = 4096

' ==================================================================
' = 概要    クリップボードにテキストを設定（Win32Apiを使用）
' = 引数    sInStr      String  [in]  設定対象文字列
' = 戻値                Boolean       設定結果
' = 覚書    Win32APIを使用する。
' =         ※ クリップボードは DataObject の PutInClipboard でも利用
' =            可能｡しかし､DataObject は参照設定が必要なうえ､特定のク
' =            リップボード形式には貼り付けされない｡（CF_UNICODETEXT
' =            のみで CF_TEXTへは貼り付けされない）
' =            上記のように DataObject を使用したくない場合に本関数
' =            を利用すること｡
' = 依存    user32/OpenClipboard()
' =         user32/EmptyClipboard()
' =         user32/CloseClipboard()
' =         user32/SetClipboardData()
' =         kernel32/GlobalAlloc()
' =         kernel32/GlobalLock()
' =         kernel32/GlobalUnlock()
' =         kernel32/lstrcpy()
' = 所属    Mng_Clipboard.bas
' ==================================================================
Public Function SetToClipboard( _
    ByVal sInStr As String _
) As Boolean
#If VBA7 Then
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hClipMemory As LongPtr
    Dim lX As LongPtr
#Else
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long
    Dim lX As Long
#End If
    Dim bResult As Boolean
    bResult = True
    
    hGlobalMemory = GlobalAlloc(GHND, LenB(sInStr) + 1)   '移動可能なグローバルメモリを割り当て
    lpGlobalMemory = GlobalLock(hGlobalMemory)          'ブロックをロックして、メモリへのfarポインタを取得
    lpGlobalMemory = lstrcpy(lpGlobalMemory, sInStr)      '文字列をグローバルメモリへコピー
    
    'メモリのロック解除
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "メモリのロックを解除できません" & vbCrLf & "処理が失敗しました"
        bResult = False
    Else
        'データをコピーするクリップボードを開く
        If OpenClipboard(0&) = 0 Then
            MsgBox "クリップボードを開くことができません" & vbCrLf & "処理が失敗しました"
            bResult = False
            Exit Function
        End If
        
        lX = EmptyClipboard()    'クリップボードの内容を消去
        hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory) 'データをクリップボードへコピー
    End If
    
    'クリップボードの状態チェック
    If CloseClipboard() = 0 Then
        MsgBox "クリップボードを閉じることができません"
        bResult = False
    End If
    SetToClipboard = bResult
End Function
    Private Function Test_SetToClipboard()
        Dim bResult As Boolean
        bResult = SetToClipboard("cliptest" & vbNewLine & "test"): Debug.Print bResult
    End Function

' ==================================================================
' = 概要    クリップボードからテキストを取得（Win32Apiを使用）
' = 引数    sOutStr     String  [Out]   取得先文字列
' = 戻値                Boolean         取得結果
' = 覚書    Win32APIを使用する。
' =         ※ クリップボードは DataObject の PutInClipboard でも利用
' =            可能｡しかし､DataObject は参照設定が必要なうえ､特定のク
' =            リップボード形式には貼り付けされない｡（CF_UNICODETEXT
' =            のみで CF_TEXTへは貼り付けされない）
' =            上記のように DataObject を使用したくない場合に本関数
' =            を利用すること｡
' = 依存    user32/OpenClipboard()
' =         user32/CloseClipboard()
' =         user32/GetClipboardData()
' =         kernel32/GlobalLock()
' =         kernel32/GlobalUnlock()
' =         kernel32/lstrcpy()
' = 所属    Mng_Clipboard.bas
' ==================================================================
Public Function GetFromClipboard( _
    ByRef sOutStr As String _
) As Boolean
#If VBA7 Then
    Dim hClipMemory As LongPtr
    Dim lpClipMemory As LongPtr
#Else
    Dim hClipMemory As Long
    Dim lpClipMemory As Long
#End If
    Dim sStr As String
    Dim lRetVal As Long
    Dim bResult As Boolean
    bResult = True
    sOutStr = ""
    
    If OpenClipboard(0&) = 0 Then
        MsgBox "クリップボードを開くことができません" & vbCrLf & "処理が失敗しました"
        bResult = False
        Exit Function
    End If
    
    ' Obtain the handle to the global memory block that is referencing the text.
    hClipMemory = GetClipboardData(CF_TEXT)
    If IsNull(hClipMemory) Then
        MsgBox "Could not allocate memory"
        bResult = False
    Else
        ' Lock Clipboard memory so we can reference the actual data string.
        lpClipMemory = GlobalLock(hClipMemory)
        
        If Not IsNull(lpClipMemory) Then
            sStr = Space$(MAXSIZE)
            Call lstrcpy(sStr, lpClipMemory)
            Call GlobalUnlock(hClipMemory)
            sStr = Mid(sStr, 1, InStr(1, sStr, Chr$(0), 0) - 1)
        Else
            MsgBox "Could not lock memory to copy string from."
            bResult = False
        End If
    End If
    
    If CloseClipboard() = 0 Then
        MsgBox "クリップボードを閉じることができません"
        bResult = False
    Else
        sOutStr = sStr
    End If
    GetFromClipboard = bResult
End Function
    Private Function Test_GetFromClipboard()
        Dim sStr As String
        Dim bResult As Boolean
        bResult = GetFromClipboard(sStr): Debug.Print bResult & ":" & sStr
    End Function

