'' push value of range into a array and no repetition
'' @range
'' @arr  : array with contain distinct value
Function pushRangeToArrayNoRepetition(ByVal Range As Excel.Range, ByRef arr() As Variant)
    Dim c As Excel.Range
    Dim str As String

    For Each c In Range
        If c.Value2 <> "" Then
            str = c.Value2
        End If

        '-1 :   no find string, -2 : array no init
        If IsInArray(str, arr) < 0 Then
            pushArray2 str, arr
        End If

    Next
End Function