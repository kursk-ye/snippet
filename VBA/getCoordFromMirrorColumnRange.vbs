''return source cell or merged cell with compare source range and target range
''''EXAMPLE
''| source  |                                                 | target      |
''| range   |                                                  | range     |
''|             |                                                 |               |
''|   ABC   |    compare value one by one      |               |
''|             |                                                |               |
''|             |                                                |   ABC     |
''|             |                                                |               |
''
''return source range and target range "ABC" cell ,
''if this cell is mered cell ,and return merged first cell and last cell
''parameter
''sRange    :   source range
''tRange    :   target range
''sCellArrr :   array,
''if searched cell is not merge, then return array has only element which is cell
''if searched cell is merged, first element return mergearea range ,second and third element
''  is merged range first cell and last cell
''tCellArr  :   array, same as above but target range
''
Sub getCoordFromMirrorColumnRange(sRange As Range, tRange As Range, ByRef sCellArr() As Range, ByRef tCellArr() As Range)
Dim sCell As Range, tCell As Range

Set sCell = sRange.Item(1)
Set tCell = tRange.Item(1)

Do Until (sCell.row > sRange.Item(sRange.Count).row)

    Do Until (tCell.row > tRange.Item(tRange.Count).row)

    If sCell.Value2 = tCell.Value2 Then
        returnMergedFirstAndLastCell sCell, sCellArr
        returnMergedFirstAndLastCell tCell, tCellArr
        Exit Sub
    End If

    Set tCell = tCell.Offset(rowOffset:=1, columnOffset:=0)
    Loop

Set sCell = sCell.Offset(rowOffset:=1, columnOffset:=0)
Loop


End Sub

''if cell is not merged ,only return this cell in arr()
''if cell is merged, the first element of arr() is the mergearea range,
''second element and third element is first cell and last cell of merged range
''@cell :   merged range
''@arr  :   return arr() contain cell
Sub returnMergedFirstAndLastCell(cell As Range, ByRef arr() As Range)
ReDim arr(2)

If cell.MergeCells Then
    Set arr(0) = cell.MergeArea
    Set arr(1) = cell.Item(1)
    Set arr(2) = cell.Item(cell.MergeArea.Count)
Else
    Set arr(0) = cell
End If

End Sub