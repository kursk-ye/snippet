'在一个字符串中查询两个字符串，并将该字符串在两者之间的字符串返回。
'@ str           需要查询的字符串
'@ startStr      查找的第一个字符串
'@ endStr        查找的第二个字符串
Function superMidStr(str As String, startStr As String, endStr As String)
    superMidStr = Mid(str, (InStr(str, startStr) + 1), (InStr(str, endStr) - (InStr(str, startStr))))
End Function