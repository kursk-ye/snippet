'根据指定的range范围内查询单元格的颜色，返回符合iColor颜色的的位置和值
'range  指定范围
'iColor 指定颜色
'multiArr 二维数组,第二维的顺序是 行号,列号,值 行号,列号,值 [[row1, col1, value1],[row2, col2, value2] ... ]
'return number 符合条件的单元格数量(考虑合并单元格的情况)，同时这个值也是二维数组实际有值的长度
Function getBitmapColor(range As range, iColor As Long, ByRef multiArr() As Variant)
Dim cellArr(2) As Variant                   '单元格的行号,列号,值组成的数组
Dim i As Long                               '单元格索引号
Dim iCell As range                          '单元格
Dim number As Integer                       '找到的符合要求的单元格数量
Dim exist As Boolean                        '是否存在的标识，true存在,false不存在


For iIndexCell = 1 To range.Count
    If range.Item(iIndexCell).Interior.Color = iColor Then
        Set iCell = range.Item(iIndexCell).MergeArea    '考虑合并单元格的情况
        cellArr(0) = iCell.row
        cellArr(1) = iCell.column
        cellArr(2) = iCell.Item(1).Value2
        
        exist = existMultiArray(multiArr, cellArr)
        If exist = False Then
            pushMultiArray multiArr, cellArr, number
            number = number + 1
        End If
    End If
Next

getBitmapColor = number


End Function

'在一个多维数组中检查一个数组是否存在
'多维数组第二维的长度为3
'如果存在返回 true，否则返回false
Function existMultiArray(ByRef multiArr() As Variant, checkedArr() As Variant)
Dim index As Integer

For index = LBound(multiArr) To UBound(multiArr)
    If multiArr(index, 0) = checkedArr(0) _
        And multiArr(index, 1) = checkedArr(1) _
        And multiArr(index, 2) = checkedArr(2) Then
            existMultiArray = True
            Exit Function
    End If
Next

existMultiArray = False

End Function

'在一个多维数组中压入一个新元素
'多维数组第二维的长度为3,新元素的维度长度也为3
'multiArr 多维数组
'pushedArr 被压入的元素,长度为3的一维数组
'lastBound 被压入的位置，从0开始
Sub pushMultiArray(ByRef multiArr() As Variant, pushedArr() As Variant, lastBound As Integer)


'ReDim multiArr(length + 1, 2) As Variant  'VBA语法此处不能使用redim 二维数组

multiArr(lastBound, 0) = pushedArr(0)
multiArr(lastBound, 1) = pushedArr(1)
multiArr(lastBound, 2) = pushedArr(2)

End Sub