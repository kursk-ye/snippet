''return a cell whic value equal special value from a Range
'' if return cell is a merged cell then return mergearea
''@search   :   string be found
''@range    :   range be found
Function getCoordSpecialValueFromRange(search As String, r As Range) As Range
Dim c As Range

For Each c In r
    If StrComp(c.Value2, search, 1) = 0 Then
        If c.MergeCells Then
            Set getCoordSpecialValueFromRange = c.MergeArea
            Exit Function
        Else
            Set getCoordSpecialValueFromRange = c
            Exit Function
        End If

    End If
Next
End Function