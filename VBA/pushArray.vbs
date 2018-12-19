'' push  value in an array
'' @arr              :  not need dim length in advance
Sub pushArray(val As Variant, ByRef resizeArr As Variant)

On Error GoTo FIRST
ReDim Preserve resizeArr(UBound(resizeArr) + 1)

resizeArr(UBound(resizeArr)) = val

Exit Sub

FIRST:
ReDim resizeArr(0)
resizeArr(0) = val

End Sub