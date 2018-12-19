'' push  arr2 to another arr1
Function pushArrayToArray(ByRef arr1 As Variant, ByVal arr2 As Variant) As Long
  Dim i As Integer

  For i = LBound(arr2) To UBound(arr2)
    pushArray arr2(i), arr1
  Next i

End Function