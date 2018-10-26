Public Function PhoneFormat(ByVal strPhoneNumber As String) As _
String

  Dim strResult As String
  Dim iLength As Integer
  Dim strExtraChar As String
  Dim strOriginal As String
  Dim iSpaceResult As Integer
  Dim i As Integer
  
  strOriginal = strPhoneNumber
      
  ' Remove any style characters from the user input
  strPhoneNumber = Replace(strPhoneNumber, ")", "")
  strPhoneNumber = Replace(strPhoneNumber, "(", "")
  strPhoneNumber = Replace(strPhoneNumber, "-", "")
  strPhoneNumber = Replace(strPhoneNumber, ".", "")
  strPhoneNumber = Replace(strPhoneNumber, Space(1), "")
      
  iLength = Len(strPhoneNumber)
  
  'convert any letters to numbers
  For i = 1 To iLength
    Mid$(strPhoneNumber, i, i) = _
        PhoneLetterToDigit(Mid$(strPhoneNumber, i, i))
  Next i
  
  ' now, if any other chars besides numbers exist, return original string to user
  For i = 1 To iLength
    Select Case Asc(Mid$(strPhoneNumber, i, i))
      Case Is < 48, Is > 57
        strResult = strOriginal
    End Select
  Next i
  
  Select Case iLength
' user entered a lot of numbers;only format the first 10
    Case Is > 11
      If Left$(strPhoneNumber, 1) = "1" Then
        strExtraChar = Mid$(strPhoneNumber, 12)
        strPhoneNumber = Mid$(strPhoneNumber, 2, 10)
      Else
        strExtraChar = Mid$(strPhoneNumber, 11)
        strPhoneNumber = Mid$(strPhoneNumber, 1, 10)
      End If
 
' if user included the number 1 before the area code.
'We drop this number
   
    Case Is = 11
      If Left$(strPhoneNumber, 1) = "1" Then
        strPhoneNumber = Mid$(strPhoneNumber, 2)
      Else
        ' check for a space character
        iSpaceResult = InStrRev(strOriginal, Space(1))
        
        If iSpaceResult = 0 Then
          ' we have no idea what they entered
          strResult = strOriginal
          GoTo Exit_Proc
        Else
          strExtraChar = Mid$(strPhoneNumber, iSpaceResult)
          strPhoneNumber = Mid$(strPhoneNumber, 1, _
             iSpaceResult - 1)
        End If
      
      End If
    
    Case Is = 10 ' area code and phone
      strPhoneNumber = strPhoneNumber
 ' user did not include an area code; add 3 spaces
         
    Case Is = 7
        strPhoneNumber = Space(3) & strPhoneNumber
 
   ' unable to figure out what the user typed
   ' must be an extentsion and not a 'real' phone number

      Case Else
         strResult = strOriginal
         GoTo Exit_Proc
  
  End Select
    
  'Add sytle characters into phone number (format)
  strResult = Format(strPhoneNumber, "\(@@@\)\ @@@\-@@@@") & _
     Space(1) & strExtraChar
 
Exit_Proc:
  PhoneFormat = strResult
    
End Function

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