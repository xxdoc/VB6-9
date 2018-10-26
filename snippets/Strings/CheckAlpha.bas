Public Function checkAlphaNumeric(strInputText As String) As Boolean
    
    Dim intCounter As Integer
    Dim strCompare As String
    Dim strInput As String
    checkAlphaNumeric = False

    For intCounter = 1 To Len(strInputText)

        strCompare = Mid$(strInputText, intCounter, 1)
        strInput = Mid$(strInputText, intCounter + 1, Len _ (strInputText))

        If strCompare Like ("[A-Z]") Or _
            strCompare Like ("[a-z]") Or _
            strCompare Like ("#") Then
            checkAlphaNumeric = True
        Else
            checkAlphaNumeric = False
            Exit Function
        End If

    Next intCounter

End Function