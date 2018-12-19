''if cell is not merged ,only return this cell in arr()
''if cell is merged, the first element of arr() is the mergearea range,
''second element and third element is first cell and last cell of merged range
''@cell :   merged range
''@arr  :   return arr() contain cell
Sub getMergedFirstAndLastCell(cell As Range, ByRef arr() As Range)
ReDim arr(2)

If cell.MergeCells Then
    Set arr(0) = cell.MergeArea
    Set arr(1) = cell.Item(1)
    Set arr(2) = cell.Item(cell.MergeArea.Count)
Else
    Set arr(0) = cell
End If

End Sub