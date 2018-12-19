''copy value of single column range to another range
''@sRange   :   source range,range cell can be merged
''@tRange   :   target range,range cell can be merged
Sub copySingleColumnRange(ByRef sRange As Range, ByRef tRange As Range)
Dim sc As Range, tc As Range

Set sc = sRange.Item(1)
Set tc = tRange.Item(1)

Do Until (sc.row > sRange.Item(sRange.Count).row) Or (tc.row > tRange.Item(tRange.Count).row)
    tc.Value2 = sc.Value2
    Set tc = tc.Offset(rowOffset:=1, columnOffset:=0)
    Set sc = sc.Offset(rowOffset:=1, columnOffset:=0)
Loop

End Sub