'时间格式是20150914 08:35:37
'检查两个时刻间隔是否小于2分钟，如果小于就标黄
Private Sub checkDuration()
Dim checkedSheets As Object
Dim checkedSheet As Excel.Worksheet
Dim bottomRowNo As Integer   '最底层行号
Dim row
Dim startTime As Integer
Dim nextTime As Integer
Dim startUsername As String
Dim nextUsername As String

Dim startDay
Dim startMM
Dim nextDay
Dim nextMM

Set checkedSheets = Workbooks(1).Worksheets

For Each checkedSheet In checkedSheets
    bottomRowNo = getBottomRow(checkedSheet, 1)
    
    For row = 2 To bottomRowNo - 1
        startUsername = checkedSheet.Cells(row, 2).Value
        nextUsername = checkedSheet.Cells(row + 1, 2).Value
        
        startTime = Mid(checkedSheet.Cells(row, 5).Value, 13, 2) '获取时间字符串，样式为20150914 08:35:37
        nextTime = Mid(checkedSheet.Cells(row + 1, 5).Value, 13, 2)
        
        'startDay 为转换后的日期 格式为2015/09/14
        startDay = Mid(Left(checkedSheet.Cells(row, 5).Value, 8), 1, 4) & "/" & Mid(Left(checkedSheet.Cells(row, 5).Value, 8), 5, 2) & "/" & Mid(Left(checkedSheet.Cells(row, 5).Value, 8), 7, 2)
        'startMM 为转换后的时分秒 格式为08:35:37
        startMM = Right(checkedSheet.Cells(row, 5).Value, 8)
        
        nextDay = Mid(Left(checkedSheet.Cells(row + 1, 5).Value, 8), 1, 4) & "/" & Mid(Left(checkedSheet.Cells(row + 1, 5).Value, 8), 5, 2) & "/" & Mid(Left(checkedSheet.Cells(row + 1, 5).Value, 8), 7, 2)
        nextMM = Right(checkedSheet.Cells(row + 1, 5).Value, 8)
        
        If (startUsername = nextUsername) Then
        		'先比较日期，如果不是同一天就不用继续比较
            If DateDiff("d", startDay, nextDay) >= 1 Then
                GoTo NEXTROW
            '再比较时间，如果超过2分钟就不标黄，否则该行标黄
            ElseIf DateDiff("n", startMM, nextMM) >= 2 Then
                GoTo NEXTROW
            Else
                checkedSheet.Rows(row + 1).Interior.ColorIndex = 6
            End If
            
        End If
        
NEXTROW:
        
    Next
    
Next

MsgBox "检查完毕"


End Sub
