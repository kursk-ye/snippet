'对制定的单元格添加批注，并标识该单元格及所在行的颜色
'@ cell          添加批注的单元格
'@ comment       所要添加的批注
'@ ifColor       是否要修改单元格的颜色，true 修改颜色,false 不修改颜色
'@ error_count   单元格计数器
Sub comment_cell(cell As Excel.Range, _
                comment As String, _
                Optional ByRef ifColor As Boolean, _
                Optional ByRef error_count As Integer)

If ifColor = True Then
    cell.Parent.Rows(cell.row).Interior.ColorIndex = 6
    cell.Interior.ColorIndex = 33
End If

If cell.comment Is Nothing Then
    cell.AddComment
End If
cell.comment.Text Text:=comment & Chr(10) & cell.comment.Text

error_count = error_count + 1

End Sub