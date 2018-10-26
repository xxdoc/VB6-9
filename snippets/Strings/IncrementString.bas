Function IncrementString(ByVal strString As String) As String
'
' Increments a string counter
' e.g.  "a" -> "b"
'       "az" -> "ba"
'       "zzz" -> "aaaa"
'
' strString is the string to increment, assumed to be lower-case alphabetic
' Return value is the incremented string
'

  Dim lngLenString As Long
  Dim strChar As String
  Dim lngI As Long
  
  lngLenString = Len(strString)
  
  ' Start at far right
  For lngI = lngLenString To 0 Step -1
  
    ' If we reach the far left then add an A and exit
    If lngI = 0 Then
      strString = "a" & strString
      Exit For
    End If
    
    ' Consider next character
    strChar = Mid(strString, lngI, 1)
    If strChar = "z" Then
      ' If we find Z then increment this to A
      ' and increment the character after this (in next loop iteration)
      strString = Left$(strString, lngI - 1) & "a" & Mid(strString, lngI + 1, lngLenString)
    Else
      ' Increment this non-Z and exit
      strString = Left$(strString, lngI - 1) & Chr(Asc(strChar) + 1) & Mid(strString, lngI + 1, lngLenString)
      Exit For
    End If
    
  Next lngI

  IncrementString = strString
  Exit Function
  
End Function