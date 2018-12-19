''determine one worksheet exist
''@shtName  :   string
''@wb       :   workbook
''return boolean
Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean

Dim sht As Worksheet

If wb Is Nothing Then Set wb = ThisWorkbook

On Error Resume Next
Set sht = wb.Sheets(shtName)
On Error GoTo 0
SheetExists = Not sht Is Nothing

End Function