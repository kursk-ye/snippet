'' Determine if Array is not Initialized
Function Initialized(val) As Boolean
On Error GoTo errHandler
    Dim i

    If Not IsArray(val) Then GoTo exitRoutine

    i = UBound(val)

    Initialized = True
exitRoutine:
    Exit Function
errHandler:
    Select Case Err.Number
        Case 9 'Subscript out of range
            GoTo exitRoutine
        Case Else
            Debug.Print Err.Number & ": " & Err.Description, _
                "Error in Initialized()"
    End Select
    Debug.Assert False
    Resume
End Function