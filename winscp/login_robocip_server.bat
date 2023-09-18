@echo off
set HOST=118.238.253.15:4022
set USER=endo
set PRIVATEKEY=\\wsl.localhost\Ubuntu-22.04\home\draemon_ash3\.ssh\id_rsa.ppk

start %MYEXEPATH_WINSCP% sftp://%USER%@%HOST% /privatekey=%PRIVATEKEY%

