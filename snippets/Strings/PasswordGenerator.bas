Public Function PasswordGenerator(ByVal lngLength As Long) _
  As String

' Description: Generate a random password of 'user input' length
' Parameters : lngLength - the length of the password to be 
                'generated
' Returns    : String    - Randomly generated password
' Created    : 08/21/1999 Andrew Ells-O'Brien 
             '(andrew@ellsobrien@msn.com)
  
On Error GoTo Err_Proc
  
 Dim iChr As Integer
 Dim c As Long
 Dim strResult As String
 Dim iAsc As String
 
 Randomize Timer

 For c = 1 To lngLength
   
   ' Randomly decide what set of ASCII chars we will use
   iAsc = Int(3 * Rnd + 1)
   
    'Randomly pick a char from the random set
   Select Case iAsc
     Case 1
       iChr = Int((Asc("Z") - Asc("A") + 1) * Rnd + Asc("A"))
     Case 2
       iChr = Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a"))
     Case 3
       iChr = Int((Asc("9") - Asc("0") + 1) * Rnd + Asc("0"))
     Case Else
       Err.Raise 20000, , "PasswordGenerator has a problem."
   End Select
   
   strResult = strResult & Chr(iChr)
 
 Next c
 
 PasswordGenerator = strResult
 
Exit_Proc:
 Exit Function
 
Err_Proc:
 MsgBox Err.Number & ": " & Err.Description, _
    vbOKOnly + vbCritical
 PasswordGenerator = vbNullString
 Resume Exit_Proc
 
End Function