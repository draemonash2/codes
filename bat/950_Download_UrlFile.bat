:: [usage]
::    950_Download_UrlFile.bat <url> [<file_name>]
::    
::    ex1) 950_Download_UrlFile.bat https://github.com/draemonash2/other/archive/master.zip
::    ex2) 950_Download_UrlFile.bat https://github.com/draemonash2/other/archive/master.zip master-codes.zip

set DOWNLOAD_URL=%1
if "%2" == "" (
    set FILE_NAME=%~n1%~x1
) else (
    set FILE_NAME=%2
)
@powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=new-object System.Net.WebClient;$d.Proxy.Credentials=[System.Net.CredentialCache]::DefaultNetworkCredentials;$d.DownloadFile('%DOWNLOAD_URL%','%FILE_NAME%')"
