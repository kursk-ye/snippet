''******************************************************************************************************************
Sub sortBubble(ByRef arr As Variant)
Dim i As Long, j As Long
Dim temp As Variant

i = 0

Do Until i > UBound(arr)

    j = UBound(arr)
    Do Until j <= i

    If arr(i) < arr(j) Then
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    End If

    j = j - 1
    Loop

i = i + 1
Loop
End Sub


''******************************************************************************************************************
''  sort a array by rank,same value of element should be same rank,from big to small
''  @sortArr    :   array by sorted
''  @rankArr    :   array used to rank
''  @arr1       :   another array,could be sort with sortArr
Sub sortArrayRank(ByRef sortArr() As Variant, ByRef rankArr() As Integer, Optional ByRef arr1 As Variant)
Dim rankNo As Long, i As Long

sortBubble sortArr, arr1

rankNo = 1

For i = 0 To UBound(sortArr)
    If i = 0 Then
        rankArr(i) = rankNo
    Else

        If sortArr(i) = sortArr(i - 1) Then
            rankArr(i) = rankNo
        Else
            rankNo = rankNo + 1
            rankArr(i) = rankNo
        End If

    End If
Next

End Sub