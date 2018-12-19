'将用户鼠标所选择的区域的单元格格式设置为百分比类型
Private Sub CommandButton1_Click()
Dim r As Excel.Range
Dim cell As Excel.Range

Set r = selection


For Each cell In r
    cell.NumberFormat = "0.00%"
    cell.Value2 = cell.Value2 & "%"
Next

End Sub