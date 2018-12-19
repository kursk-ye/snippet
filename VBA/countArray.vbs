Function countArray(ByVal arr As Variant) As Integer
Dim i As Variant
Dim count As Integer

For Each i In arr
    count = count + 1
Next

countArray = count

End Function