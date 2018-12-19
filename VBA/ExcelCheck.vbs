'check every cell in range is integer or zero
'@r :   range be checked
'@flag  : 1 -   check integer
'         2 -   check zero or positive integer
'         3 -   check natural number
Function checkInteger(r As Range, flag As Integer) As Boolean
Dim c As Range

checkInteger = False

For Each c In r
    If IsNumeric(c) Then

        If flag = 1 Then

        ElseIf flag = 2 Then

            If (c.Value2 < 0) Or (c.Value2 <> CInt(c.Value2)) Then
                c.Activate
                MsgBox "bad"
                End
            End If

        ElseIf flag = 3 Then

        End If

    Else
        c.Activate
        MsgBox "bad"
        End
    End If

Next

checkInteger = True
End Function