Public Function DelFromRight(sChars As String, _
 ByVal sLine As String)  As String
   'Removes unwanted characters from right of given string
   ' EXAMPLE
   '  MsgBox DelFromRight(" TEST", "THIS IS A TEST")
     'displays "THIS IS A"
  
  
   Dim iCount As Integer   
   Dim sChar As String

   sLine = ReverseString(sLine)
   sChars = ReverseString(sChars)
   sLine = DelFromLeft(sChars, sLine)
   DelFromRight = ReverseString(sLine)
   Exit Function

  
End Function
