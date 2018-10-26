Function AllCaps(stringToCheck As String) As Boolean

    AllCaps = StrComp(stringToCheck, UCase(stringToCheck), _
       vbBinaryCompare) = 0
End Function

Function AllLowerCase(stringToCheck As String) As Boolean

    AllLowerCase = StrComp(stringToCheck, LCase(stringToCheck), _
       vbBinaryCompare) = 0
End Function