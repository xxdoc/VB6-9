Public Function Abbreviate(ByVal strTitle As String) As String

    Dim strTmp As String
    Dim intID As Integer
    Dim strChar As String

    intID = 1
    strTmp = Trim$(strTitle)
    strChar = UCase$(Left$(strTmp, 1))

    Do While InStr(intID, strTmp, " ")
        intID = InStr(intID, strTmp, " ") + 1
        strChar = strChar & UCase$(Mid$(strTmp, intID, 1))
    Loop

    Abbreviate = strChar

End Function