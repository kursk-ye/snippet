'' return count of element of array,and element is not ""
Function countArrayNotEmpte(ByVal arr As Variant) As Integer
Dim count As Integer
Dim i As Integer

For i = LBound(arr) To UBound(arr)
    If arr(i) <> "" Then
        count = count + 1
    End If
Next i

countArrayNotEmpte = count

End Function