Public Function CountStr(ByVal StrSearch As String, ByVal StrFind As String) As Long

    ' *** 
    ' *** the length of the full string minus the length of the string with 
    ' *** the search string taken out, divided by the length of the search string
    ' *** 

    CountStr = (Len(StrSearch) - (Len(Replace(StrSearch, StrFind, "")))) / Len(StrFind)


End Function