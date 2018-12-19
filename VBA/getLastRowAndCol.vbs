Public Const MAX_TIMES = 40        '比较次数，经过times比较次数的比较得出的行号或列号才是最大值
Public Const MAX_NULL_CELL = 30     '找到一个内容的单元格后，探索性继续搜索的最多数量单元格

'返回给定worksheet有内容的单元格的最下列的行号和列号
'@return row_col是一个2字符长度的数组，下标0是行号，下标1是列号
Sub getLastRowAndCol(ws As Excel.Worksheet, Optional ByRef lastRow As long, Optional ByRef lastCol As long)
Dim row_col(2) As Integer

lastRow = getBottomRow(ws)
lastCol = getLastRightCol(ws)

End Sub

'通过MAX_TIMES次测试比较得到最底行,返回worksheet最下一行的行号
Function getBottomRow(ws As Excel.Worksheet)

Dim a As Long, b As Long, c As Long
Dim max As Long
Dim randomCol As Integer
Dim bottomRow As Long
Dim i As Long

bottomRow = 1 '初始bottomRow为1
i = 1

Do While i <= MAX_TIMES
    randomCol = Int((i + 2) * Rnd + i)

    a = getSpecialLastRow(ws, randomCol)
    b = getSpecialLastRow(ws, randomCol + 1)
    c = getSpecialLastRow(ws, randomCol + 2)

    max = max_three(a, b, c)

    If (max > bottomRow) Then
        bottomRow = max
        i = 1
    Else
        i = i + 1
    End If

Loop

getBottomRow = bottomRow

End Function

'获得某一个列的最底行,该函数思想是：遍历指定列的所有单元格，当找到一个单元格的值不为空时，将该单元格的
'行号设为最底行，然后继续向下搜索，并开始计数；如果下方还有单元格的值不为空，则将重新设置最底行的值，并将
'计数器清零，重新开始向下探索；直到计数器的值大于MAX_NULL_CELL,则返回该列的最底行数值.
Function getSpecialLastRow(ws As Excel.Worksheet, test_col As Integer)

Dim row, col As Long
Dim last_bottom As Long
Dim null_row_number As Integer


row = 1
col = test_col
null_row_number = 0

Do While True
    If ws.Cells(row, col).Text = "" And null_row_number < MAX_NULL_CELL Then
        If null_row_number = 0 Then
            last_bottom = row - 1
        End If

        null_row_number = null_row_number + 1
    ElseIf ws.Cells(row, col).Text <> "" And null_row_number > 0 Then
        last_bottom = row
        null_row_number = 0
    ElseIf null_row_number >= MAX_NULL_CELL Then
        getSpecialLastRow = last_bottom
        Exit Do
    End If

    row = row + 1
Loop

End Function

'3值比较
Function max_three(b1 As Long, b2 As Long, b3 As Long)

If b1 >= b2 Then
    If b1 >= b3 Then
        max_three = b1
    Else
        max_three = b3
    End If
ElseIf b2 >= b3 Then
    max_three = b2
Else
    max_three = b3
End If

End Function

'通过MAX_TIMES次测试比较,返回worksheet有内容单元格的最右边的列号
Function getLastRightCol(ws As Excel.Worksheet)

Dim a As Long, b As Long, c As Long
Dim max As Long
Dim randomRow As Integer
Dim rightCol As Long
Dim i As Long

rightCol = 1 '初始rightCol为1
i = 1

Do While i <= MAX_TIMES
    randomRow = Int((i + 2) * Rnd + i)

    a = getSpecialRightCol(ws, randomRow)
    b = getSpecialRightCol(ws, randomRow + 1)
    c = getSpecialRightCol(ws, randomRow + 2)

    max = max_three(a, b, c)

    If (max > rightCol) Then
        rightCol = max
        i = 1
    Else
        i = i + 1
    End If

Loop

getLastRightCol = rightCol


End Function

'指定worksheet的行号，返回该行最右边有内容的单元格的列号,与getSpecialLastRow的思路相同
'@ sheet
'@ row  指定行号
Function getSpecialRightCol(sheet As Excel.Worksheet _
                , row As Integer)
Dim lastCol As Integer
Dim cell As Excel.Range
Dim j As Integer, counter As Integer

counter = 0
lastCol = 0
j = 0

Do
    j = j + 1
    Set cell = sheet.Cells(row, j)

    If cell.Value2 <> "" Then
        lastCol = j
        counter = 0
    Else
        counter = counter + 1

        If counter > MAX_NULL_CELL Then
            getSpecialRightCol = lastCol
            Exit Function
        End If

    End If
Loop While True

End Function