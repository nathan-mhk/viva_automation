' Reference: https://serverfault.com/a/9039

Dim WinScriptHost
Set WinScriptHost = CreateObject("WScript.Shell")
WinScriptHost.Run Chr(34) & "C:\Users\User\Documents\nudam.bat" & Chr(34), 0
Set WinScriptHost = Nothing