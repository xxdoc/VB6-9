Function FormatDate(strDate As String) As String
Dim Lenght As Integer
Dim ReturnDate As String

On Error GoTo ErrHandle
If strDate = "" Then Exit Function
Lenght = Len(strDate)
If IsDate(strDate) Then
    ReturnDate = strDate
    GoTo FormatDateExit
End If
If Lenght < 4 Or Lenght > 10 Then GoTo ErrHandle
Select Case Lenght
    Case 4
        ReturnDate = Left(strDate, 1) & "/" & Mid(strDate, 2, 1) & "/" & Right(strDate, 2)
    Case 5
        If Not IsNumeric(strDate) Then GoTo ErrHandle
        If Left(strDate, 2) < 13 Then
           ReturnDate = Left(strDate, 2) & "/" & Mid(strDate, 3, 1) & "/" & Right(strDate, 2)
        Else
            ReturnDate = Left(strDate, 1) & "/" & Mid(strDate, 2, 2) & "/" & Right(strDate, 2)
        End If
    Case 6
        If Not IsNumeric(strDate) Then
            GoTo ErrHandle
        ElseIf Left(strDate, 2) < 13 Then
            ReturnDate = Left(strDate, 2) & "/" & Mid(strDate, 3, 2) & "/" & Right(strDate, 2)
        Else
            ReturnDate = Left(strDate, 1) & "/" & Mid(strDate, 2, 1) & "/" & Right(strDate, 4)
        End If
    Case 7
        If Not IsNumeric(strDate) Then
            GoTo ErrHandle
        ElseIf Left(strDate, 2) < 13 Then
            ReturnDate = Left(strDate, 2) & "/" & Mid(strDate, 3, 1) & "/" & Right(strDate, 4)
        Else
            ReturnDate = Left(strDate, 1) & "/" & Mid(strDate, 2, 2) & "/" & Right(strDate, 4)
        End If
    Case 8
        If Not IsNumeric(strDate) Then
            GoTo ErrHandle
        Else
            ReturnDate = Left(strDate, 2) & "/" & Mid(strDate, 3, 2) & "/" & Right(strDate, 4)
        End If
End Select
FormatDateExit:
If IsDate(ReturnDate) Then
    FormatDate = Format(ReturnDate, "MM/DD/YY")
Else
    GoTo ErrHandle
End If
Exit Function
ErrHandle:
Err.Raise 30000, "Format Date", "Not a valid Date"
End Function