Public Function FirstCharWordCapital(strSentence As String)

    Dim i As Integer
    Dim strTemp As String
    Dim intLocation As Integer
    Dim strChar As String * 1

    strTemp$ = ""

    For i% = 1 To Len(strSentence)

        strChar = Chr(Asc(Mid(strSentence, i%, 1)))
        If Len(Trim(strChar)) < 1 Then
            intLocation% = i% + 1
        End If

        If i% = intLocation% Or i% = 1 Then
            strTemp$ = strTemp$ + UCase(Chr(Asc(Mid(strSentence, i%, 1))))
        Else
            strTemp$ = strTemp$ + LCase(Chr(Asc(Mid(strSentence, i%, 1))))
        End If

    Next i

    FirstCharWordCapital = strTemp$

End Function