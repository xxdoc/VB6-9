Public Function IsAlphaNumeric(sChr As String) As Boolean
    IsAlphaNumeric = sChr Like "[0-9A-Za-z]"
End Function