'' 将一个array中的值填充到一个range中
'' @range
'' @arr  :
Function pushArrayToRange(ByRef arr As Variant, ByVal range As Excel.range)
Dim i As Long
Dim j As Long

j = 1

For i = LBound(arr) To UBound(arr)
    If  j <= range.Count Then
        range.Item(j) = arr(i)
        j = j + 1
    End If
Next i

End Function