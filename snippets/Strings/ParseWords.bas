Public Function parse(ByVal inString, Optional ByVal delimiters)
    'Take a string, and return it as a one dimensional array
    ' of individual values as delimited by any of several
    ' characters. None of those characters are returned in 
    ' the result. Provide a default list of delimiters, which
    ' should come from registry. But allow override.

    Dim delimitList, oneChar, aWord, codeCount
    Dim arrayCodes()

    If IsMissing(delimiters) Then
        'We should get these from Registry
        delimitList = " ,/!|"               
'Characters recognized as delimiters

    Else
        delimitList = delimiters            
'user can override if needed
    End If
    Dim i, j, k
    i = Len(inString)
    For j = 1 To i                         
'Read one character at a time
        
        oneChar = VBA.Strings.Mid(inString, j, 1)
        k = InStr(delimitList, oneChar)    
'Is this one a delimiter?
        If k = 0 Then
            aWord = aWord & oneChar         
'If is isn't, add to the current word
        End If
        If k <> 0 Or j = i Then             
'If it is, or if we're finished
            If aWord > "" Then
                codeCount = codeCount + 1
                ReDim Preserve arrayCodes(codeCount)
                arrayCodes(codeCount) = aWord       
'Save new word
                aWord = ""
            End If
        End If
    Next j
    parse = arrayCodes                              
'Return the array
End Function