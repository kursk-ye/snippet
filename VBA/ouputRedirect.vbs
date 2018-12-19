Set objShell = WScript.CreateObject("WScript.Shell")
strRun = "%comspec% /c ipconfig.exe > " & AddQuotes("d:\ipconfig.txt")
objShell.Run strRun, 1, True

Function AddQuotes(strInput)
AddQuotes = Chr(34) & strInput & Chr(34)
End Function