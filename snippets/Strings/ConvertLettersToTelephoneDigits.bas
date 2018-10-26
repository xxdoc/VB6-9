Function PhoneLetterToDigit(ByVal strPhoneLetter As String) As _
String
  
  Dim intDigit As Integer
  
  intDigit = Asc(UCase$(strPhoneLetter))
    
  If intDigit >= 65 And intDigit <= 90 Then

    If intDigit = 81 Or 90 Then ' Q or Z
      intDigit = intDigit - 1
    End If

    intDigit = (((intDigit - 65) \ 3) + 2)
    PhoneLetterToDigit = intDigit
  Else
    PhoneLetterToDigit = strPhoneLetter
  End If

End Function