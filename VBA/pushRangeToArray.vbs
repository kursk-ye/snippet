'' push value of range into a array
'' @range
'' @arr  : array with contain distinct value
Function pushRangeToArray(ByVal Range As Excel.Range, ByRef arr() As Variant)
    Dim c As Excel.Range
    Dim str As String

    For Each c In Range
        If c.Value2 <> "" Then
            str = c.Value2
        End If

        pushArray2 str, arr

    Next
End Function