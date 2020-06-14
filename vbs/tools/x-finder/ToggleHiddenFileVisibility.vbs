'Option Explicit
'Const EXECUTION_MODE = 255 '0:Explorerから実行、1:X-Finderから実行、other:デバッグ実行

'####################################################################
'### 設定
'####################################################################


'####################################################################
'### 本処理
'####################################################################
Const sPROG_NAME = "隠しファイル表示切り替え"

Dim bIsContinue
bIsContinue = True

If bIsContinue = True Then
    If EXECUTION_MODE = 1 Then 'X-Finderから実行
        'Do Nothing
    Else
        MsgBox "このスクリプトはX-Finder以外では実行できません。", vbOKOnly, sPROG_NAME
        MsgBox "処理を中断します。", vbOKOnly, sPROG_NAME
        bIsContinue = False
    End If
Else
    'Do Nothing
End If

If bIsContinue = True Then
    If InStr( WScript.Env("Style"), "h" ) > 0 Then
        MsgBox "隠しファイル、システムファイルを【非表示】にします。", vbOKOnly, sPROG_NAME
    Else
        MsgBox _
            "隠しファイル、システムファイルを【表示】します。" & vbNewLine & _
            "" & vbNewLine & _
            "(※) エクスプローラーのフォルダ設定にて「保護されたオペレーティングシステムファイルを表示しない（推奨）」がチェックされている場合、システムファイルは表示されません。" _
            , vbOKOnly, sPROG_NAME
    End If
    WScript.Exec("Change:Style ~h")
    WScript.Exec("Refresh:4")
Else
    'Do Nothing
End If
