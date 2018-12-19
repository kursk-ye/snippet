'' Determines whether a string has the same value in an array
'' @stringToBeFound  :  to be found string
'' @arr              :  use to found in array
Function IsInArray(stringToBeFound As String, arr As Variant) As Long
Dim i As Long
' default return value if value not found in array
IsInArray = -1

if Initialized(arr) then
    For i = LBound(arr) To UBound(arr)
      If StrComp(stringToBeFound, arr(i), vbTextCompare) = 0 Then
        IsInArray = i
        Exit For
      End If
    Next i
else
    IsInArray = -2
End If

End Function