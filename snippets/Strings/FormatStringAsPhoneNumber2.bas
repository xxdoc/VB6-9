'Takes any entered phone number and returns it in ###-#### format
'or (###) ###-####

Public Function FormatPhoneNumber(ByVal sNumToBeFormatted As _
   String) As String

Dim iNumberLength As Integer 'Used for the Phone Number length
   
'Trim any leading and trailing spaces

sNumToBeFormatted = Trim$(sNumToBeFormatted)
   
'Length of the phone number.

iNumberLength = Len(sNumToBeFormatted)
   
Select Case iNumberLength

  Case 7  'Format : #######

    FormatPhoneNumber = Left$(sNumToBeFormatted, 3) & _
        "-" & Right$(sNumToBeFormatted, 4)
    Exit Function

  Case 8  'Format : ###-#### or ### ####

    If Mid$(sNumToBeFormatted, 4, 1) = "-" Then
       FormatPhoneNumber = sNumToBeFormatted
       Exit Function
    Else
       FormatPhoneNumber = Left$(sNumToBeFormatted, 3) & "-" & _
          Right$(sNumToBeFormatted, 4)
       Exit Function
    End If

  Case 10 'Format : ##########

 FormatPhoneNumber = "(" & Left$(sNumToBeFormatted, 3) & ") " _
   & Mid$(sNumToBeFormatted, 4, 3) & "-" & _
     Right$(sNumToBeFormatted, 4)
 
   Exit Function

  Case 11 'Format ######-####

 FormatPhoneNumber = "(" & Left$(sNumToBeFormatted, 3) & ") " & _
       Right$(sNumToBeFormatted, 8)
    Exit Function

  Case 12 'Format : ### ###-####

 FormatPhoneNumber = "(" & Left$(sNumToBeFormatted, 3) & ") " & _
      Mid$(sNumToBeFormatted, 5, 3) & "-" & _
      Right$(sNumToBeFormatted, 4)
    Exit Function

  Case 13 'Format : (###)###-####
     FormatPhoneNumber = Left(sNumToBeFormatted, 5) & " " & _
        Right(sNumToBeFormatted, 8)
     Exit Function


  Case Else
        'Return Value Passed
     FormatPhoneNumber = sNumToBeFormatted
           
End Select

End Function