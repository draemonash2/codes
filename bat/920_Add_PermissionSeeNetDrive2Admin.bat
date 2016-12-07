::管理者権限を持つプログラムからネットワークドライブを参照できるようにする。
::本バッチファイル実行後、PC を再起動すること。

@echo off
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v EnableLinkedConnections /t REG_DWORD /d 1 /f
pause
