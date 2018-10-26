Public Function IsValidSSN(ByVal SSN As String) As Boolean

    'Determines if SSN is a valid social security number
    'requires SSN to be in either "#########" or "###-##-####" format

    IsValidSSN = (SSN Like "###-##-####") Or SSN Like ("#########")

End Function