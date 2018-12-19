'' find splite sign in string,return array.if not find return 0
Function supSplit(ByVal str As String, Optional ByRef flag As Integer) As String()
Dim aStr(0) As String
flag = 1

If InStr(str, ",") > 0 Then
    supSplit = Split(str, ",")
ElseIf InStr(str, ":") > 0 Then
    supSplit = Split(str, ":")
ElseIf InStr(str, "|") > 0 Then
    supSplit = Split(str, "|")
ElseIf InStr(str, "\") > 0 Then
    supSplit = Split(str, "\")
ElseIf InStr(str, ";") > 0 Then
    supSplit = Split(str, ";")
ElseIf InStr(str, "_") > 0 Then
    supSplit = Split(str, "_")
ElseIf InStr(str, "-") > 0 Then
    supSplit = Split(str, "-")
Else
    aStr(0) = str
    supSplit = aStr
    flag = 0
End If


End Function