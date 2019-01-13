Attribute VB_Name = "modStringHandling"
'***************************************************************************************
'***
'***    FUNCTION LIST:
'***
'***        Abbreviate              IsAlphabetical
'***        CharCount               IsAlphaNumeric
'***        FormatDate              IsNumericOnly
'***        AllCaps                 ReverseString
'***        AllLowerCase            Post
'***        FindNext                ParameterCount
'***        FlipCase                ParameterValue
'***        FormatPhoneNumber       NTC
'***        TextFromFile            GetRandomString
'***        RemoveChars
'***
'***************************************************************************************

'*************** Begin Definitions for Post function ***********************************
Const MAXCHARS = 5000
'*************** End Definitions for Post function *************************************

'*************** Begin Definitions for Number to Text function *************************
Public one, two, three, four, five, six, no As String
Dim ar(20) As String
Dim A(10) As String
Dim n1 As Double
'*************** End Definitions for Number to Text function ***************************

Public Function Abbreviate(ByVal strTitle As String) As String
'***************************************************************************************
'***
'***    FUNCTION:       Abbreviate
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Abbreviates a string, based on the first letters of each word
'***    RETURNS:        String, abbreviated
'***    USAGE:          ret = Abbreviate("Microsoft Visual Basic")
'***                        returns "MVB"
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
Dim strTmp As String
Dim intID As Integer
Dim strChar As String

    intID = 1
    strTmp = Trim$(strTitle)
    strChar = UCase$(Left$(strTmp, 1))
    
    Do While InStr(intID, strTmp, " ")
        intID = InStr(intID, strTmp, " ") + 1
        strChar = strChar & UCase$(Mid$(strTmp, intID, 1))
    Loop
    Abbreviate = strChar

End Function
Public Function CharCount(OrigString As String, _
  Chars As String, Optional CaseSensitive As Boolean = False) _
  As Long
'***************************************************************************************
'***
'***    FUNCTION:       CharCount
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Counts the number of occurences of a character or sequence of
'***                    characters within a string
'***    RETURNS:        Number of Occurrences of Chars in OrigString
'***    SIDE EFFECTS:   Useful in all versions of VB
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
Dim lLen As Long
Dim lCharLen As Long
Dim lAns As Long
Dim sInput As String
Dim sChar As String
Dim lCtr As Long
Dim lEndOfLoop As Long
Dim bytCompareType As Byte

sInput = OrigString
If sInput = "" Then Exit Function
lLen = Len(sInput)
lCharLen = Len(Chars)
lEndOfLoop = (lLen - lCharLen) + 1
bytCompareType = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then lAns = lAns + 1
    Next

CharCount = lAns

End Function
Function FormatDate(strDate As String) As String
'***************************************************************************************
'***
'***    FUNCTION:       FormatDate
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        input mask for dates take in a string with or without
'***                    formatters (/ or -) and formats it to ##/##/####
'***    RETURNS:        formatted date string
'***    USAGE:          Debug.Print formatdate("5-28-1970") returns "05/28/1970"
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
Dim Lenght As Integer
Dim ReturnDate As String

On Error GoTo ErrHandle
If strDate = "" Then Exit Function
Lenght = Len(strDate)
If IsDate(strDate) Then
    ReturnDate = strDate
    GoTo FormatDateExit
End If
If Lenght < 4 Or Lenght > 10 Then GoTo ErrHandle
Select Case Lenght
    Case 4
        ReturnDate = Left(strDate, 1) & "/" & Mid(strDate, 2, 1) & "/" & Right(strDate, 2)
    Case 5
        If Not IsNumeric(strDate) Then GoTo ErrHandle
        If Left(strDate, 2) < 13 Then
           ReturnDate = Left(strDate, 2) & "/" & Mid(strDate, 3, 1) & "/" & Right(strDate, 2)
        Else
            ReturnDate = Left(strDate, 1) & "/" & Mid(strDate, 2, 2) & "/" & Right(strDate, 2)
        End If
    Case 6
        If Not IsNumeric(strDate) Then
            GoTo ErrHandle
        ElseIf Left(strDate, 2) < 13 Then
            ReturnDate = Left(strDate, 2) & "/" & Mid(strDate, 3, 2) & "/" & Right(strDate, 2)
        Else
            ReturnDate = Left(strDate, 1) & "/" & Mid(strDate, 2, 1) & "/" & Right(strDate, 4)
        End If
    Case 7
        If Not IsNumeric(strDate) Then
            GoTo ErrHandle
        ElseIf Left(strDate, 2) < 13 Then
            ReturnDate = Left(strDate, 2) & "/" & Mid(strDate, 3, 1) & "/" & Right(strDate, 4)
        Else
            ReturnDate = Left(strDate, 1) & "/" & Mid(strDate, 2, 2) & "/" & Right(strDate, 4)
        End If
    Case 8
        If Not IsNumeric(strDate) Then
            GoTo ErrHandle
        Else
            ReturnDate = Left(strDate, 2) & "/" & Mid(strDate, 3, 2) & "/" & Right(strDate, 4)
        End If
End Select
FormatDateExit:
If IsDate(ReturnDate) Then
    FormatDate = Format(ReturnDate, "MM/DD/YY")
Else
    GoTo ErrHandle
End If
Exit Function
ErrHandle:
Err.Raise 30000, "Format Date", "Not a valid Date"
End Function
Function AllCaps(stringToCheck As String) As Boolean
'***************************************************************************************
'***
'***    FUNCTION:       AllCaps
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Determines if a string is all uppercase
'***    RETURNS:        boolean true/false
'***    USAGE:          Debug.Print AllCaps("TEST STRING") returns true
'***                    Debug.Print AllCaps("Test String") returns false
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************

AllCaps = StrComp(stringToCheck, UCase(stringToCheck), vbBinaryCompare) = 0

End Function
Function AllLowerCase(stringToCheck As String) As Boolean
'***************************************************************************************
'***
'***    FUNCTION:       AllLowerCase
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Determines if a string is all lower case
'***    RETURNS:        boolean true/false
'***    USAGE:          Debug.Print AllLowerCase("test string") returns true
'***                    Debug.Print AllLowerCase("Test String") returns false
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************

AllLowerCase = StrComp(stringToCheck, LCase(stringToCheck), vbBinaryCompare) = 0

End Function
Public Function FindNext(ByVal text As String, ByVal Search As String, _
 Optional ByVal Position As Long, Optional CaseSensitive As Boolean, _
 Optional ByVal Up As Boolean) As Long
'***************************************************************************************
'***
'***    FUNCTION:       FindNext
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Finds a string within a string, up or down, case sensitive
'***                    or case insensitive.
'***    RETURNS:        Positon the string was found
'***    USAGE:          debug.print findnext("This is a string","is",0,false,false)
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
Dim lPos As Long
Dim lFind As Long
    
If text <> "" Then
    
    If Position < 1 Then Position = 1
    
    If Position > Len(text) Then Position = Len(text)
    
    If Up Then
        
        lPos = Position - 1
        While lPos > 0
            
            lFind = InStr(lPos, text, Search, Abs(Not CaseSensitive))
            If lFind = lPos Then
                    
                lPos = 0
                
            Else
                    
                lFind = 0
                lPos = lPos - 1
                
            End If
        
        Wend
            
        FindNext = lFind
        
    Else
        
        FindNext = InStr(Position + 1, text, Search, _
        Abs(Not CaseSensitive))
    
    End If

End If

End Function
Public Function FlipCase(ByVal ThisText As String) As String
'***************************************************************************************
'***
'***    FUNCTION:       FlipCase
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Reverses the case of all characters in a string
'***    RETURNS:        Flipped case string
'***    USAGE:          debug.print FlipCase("Hello") returns "hELLO"
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
Dim Letter As String, Temp As String
Dim n As Long, nLen As Long
    
nLen = Len(ThisText)

For n = 1 To nLen
    
    Letter = Mid(ThisText, n, 1)

    If Asc(Letter) > 96 And Asc(Letter) < 123 Then
        
        Letter = UCase(Letter)
    
    ElseIf Asc(Letter) > 64 And Asc(Letter) < 91 Then
        
        Letter = LCase(Letter)
    
    End If
    
    Temp = Temp & Letter

Next

FlipCase = Temp

End Function
Public Function FormatPhoneNumber(ByVal sNumToBeFormatted As String) As String
'***************************************************************************************
'***
'***    FUNCTION:       FormatPhoneNumber
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Formats a phone number to ###-####
'***    RETURNS:        string containing formatted phone number
'***    USAGE:          debug.print FormatPhoneNumber("3605551212") returns "(360) 555-1212"
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************

Dim iNumberLength As Integer

sNumToBeFormatted = Trim$(sNumToBeFormatted)
   
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
Public Function TextFromFile(fInStream As String) As String
'***************************************************************************************
'***
'***    FUNCTION:       TextFromFile
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Reads text from a file quickly.  Much quicker on large files
'***                    than the usual While...Wend loop.
'***    RETURNS:        text from file
'***    USAGE:          debug.print TextFromFile("C:\Windows\Programs.txt")
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
Dim i As Long, strText As String

i = FreeFile
strText = ""

Open fInStream For Input Lock Write As #i
    
    Screen.MousePointer = vbHourglass
    DoEvents
    strText = StrConv(InputB$(LOF(i), i), vbUnicode)

Close #i

Screen.MousePointer = vbDefault
TextFromFile = strText

End Function
Public Function RemoveChars(ByVal pMessage As String, pRemovable As String) As String
'***************************************************************************************
'***
'***    FUNCTION:       RemoveChars
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Remove characters from a string.  Intended for use with VB5 or
'***                    below.  If using VB6 or higher, use the built-in Replace function.
'***    RETURNS:        string with characters removed
'***    USAGE:          debug.print RemoveChars("abacadae","a") returns "bcde"
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
Dim lMessage As String
Dim lCurChar As String
Dim X As Integer
    
For X = 1 To Len(pMessage)

    lCurChar = Mid(pMessage, X, 1)
    If InStr(pRemovable, lCurChar) = 0 Then lMessage = lMessage & lCurChar

Next X

RemoveChars = lMessage

End Function
Public Function IsAlphaBetical(TestString As String) As Boolean
'***************************************************************************************
'***
'***    FUNCTION:       IsAlphaBetical
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Tests to see if a string contains only alphabetical characters
'***    RETURNS:        Boolean true/false
'***    USAGE:          debug.print IsAlphaBetical("Abcde") returns true
'***                    debug.print IsAlphaBetical("Abc12") returns false
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
    
Dim sTemp As String
Dim iLen As Integer
Dim iCtr As Integer
Dim sChar As String
    
sTemp = TestString
iLen = Len(sTemp)

If iLen > 0 Then
    
    For iCtr = 1 To iLen
        
        sChar = Mid(sTemp, iCtr, 1)
        If Not sChar Like "[A-Za-z]" Then Exit Function
    
    Next
    
    IsAlphaBetical = True

End If
    
End Function
Public Function IsAlphaNumeric(TestString As String) As Boolean
'***************************************************************************************
'***
'***    FUNCTION:       IsAlphaNumeric
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Tests to see if a string contains only alphabetical and Numeric
'***                    characters
'***    RETURNS:        Boolean true/false
'***    USAGE:          debug.print IsAlphaNumeric("Abcde") returns false
'***                    debug.print IsAlphaNumeric("Abc12") returns true
'***                    debug.print IsAlphaNumeric("A-23")  returns false
'***                    debug.print IsAlphaNumeric("1234")  returns false
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************

Dim sTemp As String
Dim iLen As Integer
Dim iCtr As Integer
Dim sChar As String
    
sTemp = TestString
iLen = Len(sTemp)

If iLen > 0 Then
        
    For iCtr = 1 To iLen
        
        sChar = Mid(sTemp, iCtr, 1)
        If Not sChar Like "[0-9A-Za-z]" Then Exit Function
    
    Next
    
    IsAlphaNumeric = True
    
End If
    
End Function
Public Function IsNumericOnly(TestString As String) As Boolean
'***************************************************************************************
'***
'***    FUNCTION:       IsNumericOnly
'***    DATE:           08/27/2003
'***
'***    PURPOSE:        Tests to see if a string contains only Numeric Characters
'***    RETURNS:        Boolean true/false
'***    USAGE:          debug.print IsNumericOnly("Abcde") returns false
'***                    debug.print IsNumericOnly("Abc12") returns false
'***                    debug.print IsNumericOnly("12345") returns true
'***                    debug.print IsNumericOnly("99.99") returns false
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/27/2003  Initial Creation
'***
'***************************************************************************************
    
Dim sTemp As String
Dim iLen As Integer
Dim iCtr As Integer
Dim sChar As String
    
sTemp = TestString
iLen = Len(sTemp)
    
If iLen > 0 Then
        
    For iCtr = 1 To iLen
            
        sChar = Mid(sTemp, iCtr, 1)
        If Not sChar Like "[0-9]" Then Exit Function
        
    Next
    
    IsNumericOnly = True
    
End If
    
End Function
Public Function ReverseString(TextToReverse As String) As String
'***************************************************************************************
'***
'***    FUNCTION:       ReverseString

'***    DATE:           08/28/2003

'***
'***    PURPOSE:        Reverses a string of text
'***    RETURNS:        Reversed String
'***    USAGE:          Text1.Text = ReverseString(Text1.Text)
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    08/28/2003  Initial Creation
'***
'***************************************************************************************
Dim Kount As Long
Dim Work As String

For Kount = Len(TextToReverse) To 1 Step -1

    Work = Work & Mid(TextToReverse, Kount, 1)
    
Next

ReverseString = Work

End Function
Public Function ParameterValue(ParseCharacter As String, _
                               tString As Variant, _
                               Index As Integer) As String
'***************************************************************************************
'***
'***    FUNCTION:       ParameterValue
'***    DATE:           09/26/2002
'***
'***    PURPOSE:        Returns a field value from a delimited string, given the delimiter,
'***                    string, and which field to pick
'***    RETURNS:        string value of field in the string
'***    USAGE:          outstring = ParameterValue(",",instring,3)
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    07/30/2003  Initial Creation
'***    09/25/2003  Addition to main string handling module
'***
'***************************************************************************************

Dim CurrentPosition As Integer
Dim ParseToPosition As Integer
Dim CurrentToken As Integer
Dim TempString As String

TempString = Trim(tString) + ParseCharacter

If Len(TempString) = 1 Then Exit Function

CurrentPosition = 1
CurrentToken = 1

Do
    ParseToPosition = InStr(CurrentPosition, TempString, _
        ParseCharacter)
    
    If Index = CurrentToken Then
        
        ParameterValue = Mid$(TempString, CurrentPosition, _
            ParseToPosition - CurrentPosition)
        Exit Function

    End If

    CurrentToken = CurrentToken + 1
    CurrentPosition = ParseToPosition + 1

Loop Until (CurrentPosition >= Len(TempString))

End Function
Public Function ParameterCount(ParseCharacter As String, _
                               tString As Variant) As Integer
'***************************************************************************************
'***
'***    FUNCTION:       ParameterCount
'***    DATE:           09/26/2002
'***
'***    PURPOSE:        Counts the number of fields in a delimited string, given the
'***                    the delimiter and the string
'***    RETURNS:        Number of fields in the string
'***    USAGE:          ret = ParameterCount(",",txtString)
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    07/30/2003  Initial Creation
'***    09/25/2003  Addition to main string handling module
'***
'***************************************************************************************
          
Dim CurrentPosition As Integer
Dim ParseToPosition As Integer
Dim CurrentToken As Integer
Dim TempString As String

TempString = Trim(tString) + ParseCharacter
  
If Len(TempString) = 1 Then Exit Function
  
CurrentPosition = 1
CurrentToken = 1
  
Do
    ParseToPosition = InStr(CurrentPosition, TempString, ParseCharacter)
    CurrentToken = CurrentToken + 1
    CurrentPosition = ParseToPosition + 1
  
Loop Until (CurrentPosition >= Len(TempString))
  
  ParameterCount = CurrentToken - 1

End Function
Public Sub Post(tbxEditBox As TextBox, sNewText As String)
'***************************************************************************************
'***
'***    SUBROUTINE:     Post
'***    DATE:           07/30/2003
'***
'***    PURPOSE:        Makes a multi-line textbox into a scrolling textbox, chat-style.
'***    RETURNS:        None
'***    USAGE:          Post Text1, "Text to Post"
'***
'***    SIDE EFFECTS:   Constant MAXCHARS must be declared at the top of the module.  Do
'***                    not exceed 50000 for MAXCHARS.
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    07/30/2003  Initial Creation
'***    09/25/2003  Addition to main string handling module
'***
'***************************************************************************************
sNewText = sNewText & vbCrLf
    
With tbxEditBox
    
    If Len(sNewText) + Len(.text) > MAXCHARS Then
        
        'Scroll some text off the top to make more room
        .text = Mid$(.text, InStr(100 + Len(sNewText), .text, vbCrLf) + 2)
    
    End If
    
    .SelStart = Len(.text)
    .SelText = sNewText

End With

End Sub
Public Function ntc(num As Double) As String
'***************************************************************************************
'***
'***    SUBROUTINE:     ntc
'***    DATE:           07/30/2003
'***
'***    PURPOSE:        Converts a number to its text (written) equivalent
'***    RETURNS:        None
'***    USAGE:          outstring = ntc(543)
'***                    returns "Five Hundred Forty Three"
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    09/25/2003  Initial Creation
'***
'***************************************************************************************

    ar(1) = "One"
    ar(2) = "Two"
    ar(3) = "Three"
    ar(4) = "Four"
    ar(5) = "Five"
    ar(6) = "Six"
    ar(7) = "Seven"
    ar(8) = "Eight"
    ar(9) = "Nine"
    ar(10) = "Ten"
    ar(11) = "Eleven"
    ar(12) = "Twelve"
    ar(13) = "Thirteen"
    ar(14) = "Fourteen"
    ar(15) = "Fifteen"
    ar(16) = "Sixteen"
    ar(17) = "Seventeen"
    ar(18) = "Eighteen"
    ar(19) = "Nineteen"
    ar(20) = "Twenty"
    A(2) = "Twenty"
    A(3) = "Thirty"
    A(4) = "Fourty"
    A(5) = "Fifty"
    A(6) = "Sixty"
    A(7) = "Seventy"
    A(8) = "Eighty"
    A(9) = "Ninety"
    one = ""
    two = ""
    three = ""
    four = ""
    five = ""
    six = ""
    Count = 0
    ln = Len(Trim(Str(num)))


    For i = 0 To (ln - 1) Step 1
        n1 = Val(Right(Trim(Str(num)), i + 1))
        no = Trim(Str(n1))


        If n1 <= 9 And n1 >= 0 Then
            one = onecheck(n1)
        Else


            If n1 <= 99 Then
                two = tencheck(n1)
            Else


                If n1 <= 999 Then
                    three = huncheck(n1)
                Else


                    If n1 <= 999999 Then
                        four = thcheck(n1, 1000, 9999, 10000, 19999, 20000, 99999, 100000, 999999, " Thousand")
                    Else


                        If n1 <= 999999999 Then
                            five = thcheck(n1, 1000000, 9999999, 10000000, 19999999, 20000000, 99999999, 100000000, 999999999, " Million")
                        Else
                            six = thcheck(n1, 1000000000, 9999999999#, 10000000000#, 19999999999#, 20000000000#, 99999999999#, 100000000000#, 999999999999#, " Billion")
                        End If

                    End If

                End If

            End If

        End If


        If n1 > 9 And one = "Zero" Then
            one = ""
        End If

    Next i

    stng = six & " " & five & " " & four & " " & three & " " & two & " " & one
    ntc = Trim(stng)
    p1 = InStr(2, ntc, " and")
    p2 = InStrRev(ntc, " and")


    If p1 > 0 And p2 > 0 Then
        ntc = Mid(ntc, 1, p1) + Mid(ntc, p2)


        If Right(ntc, 4) = " and" Then
            ntc = Mid(ntc, 1, InStrRev(ntc, " and"))
        End If

    End If

End Function
Function onecheck(n As Double) As String
'***************************************************************************************
'***
'***    NTC SUBFUNCTION
'***    Ones Check
'***
'***************************************************************************************
    Count = Count + 1


    If n <> 0 Then
        ones = ar(n1)


        If Count > 1 And one = "Zero" Then
            ones = ""
        End If

    Else
        ones = "Zero"
    End If

    onecheck = ones
End Function
Function tencheck(n As Double) As String
'***************************************************************************************
'***
'***    NTC SUBFUNCTION
'***    Tens Check
'***
'***************************************************************************************


    If n >= 10 And n <= 20 Then
        tens = ar(n1)
        one = ""
    Else
        two = Left(no, 1)
        tens = A(Val(two))


        If one = "Zero" Then
            one = ""
        End If

    End If

    tencheck = tens
End Function
Function huncheck(n As Double) As String
'***************************************************************************************
'***
'***    NTC SUBFUNCTION
'***    Hundreds Check
'***
'***************************************************************************************

    ps = Left(no, 1)
    hun = ar(Val(ps)) + " Hundred"


    If two <> "" Then
        two = "and " + two
    Else


        If one = "Zero" And two = "" Then
            one = ""
        Else


            If one <> "Zero" And two = "" Then
                one = "and " + one
            End If

        End If

    End If

    huncheck = hun
End Function
Function thcheck(n As Double, n11 As Double, n2 As Double, n3 As Double, n4 As Double, n5 As Double, n6 As Double, n7 As Double, n8 As Double, text As String) As String
'***************************************************************************************
'***
'***    NTC SUBFUNCTION
'***    Thousands Check
'***
'***************************************************************************************


    If n >= n11 And n <= n2 Then
        ps = Left(no, 1)
        th = ar(Val(ps)) + text
    Else


        If n >= n3 And n <= n4 Then
            ps = Left(no, 2)
            th = ar(Val(ps)) + text
        Else


            If n >= n5 And n <= n6 Then
                ps = Left(no, 1)
                pp = Mid(no, 2, 1)
                tt = " "


                If pp <> "0" Then
                    tt = ar(Val(pp))
                    th = A(Val(ps)) + " " + tt + text
                Else
                    th = A(Val(ps)) + text
                End If

            Else


                If n >= n7 And n <= n8 Then
                    hs = Left(no, 1)
                    ps = Mid(no, 2, 1)
                    pp = Mid(no, 3, 1)
                    tn = Val(Mid(no, 2, 2))
                    tt = " "


                    If tn >= 10 And tn <= 20 Then
                        th = ar(Val(hs)) + " Hundred " + ar(Val(tn)) + " " + text
                    Else


                        If pp <> "0" And ps <> "0" Then
                            tt = ar(Val(pp))
                            th = ar(Val(hs)) + " Hundred " + A(Val(ps)) + " " + tt + text
                        ElseIf pp = "0" And ps = "0" Then
                            th = ar(Val(hs)) + " Hundred " + text
                        ElseIf pp = 0 Then
                            th = ar(Val(hs)) + " Hundred " + A(Val(ps)) + text
                        ElseIf ps = "0" Then
                            tt = ar(Val(pp))
                            th = ar(Val(hs)) + " Hundred " + tt + text
                        End If

                    End If

                End If

            End If

        End If

    End If

    thcheck = th
End Function

Public Function GetRandomString(ByVal lngLength As Long, ByVal strStringList As String) As String
'***************************************************************************************
'***
'***    SUBROUTINE:     GetRandomString
'***    DATE:           09/25/2003
'***
'***    PURPOSE:        Creates a string of random characters, of given length, within the
'***                    limits of the given stringset
'***    RETURNS:        string
'***    USAGE:          outstring = getrandomstring(10,"ABCDEFGHIJKLMNOPQRSTUVWXYZ")
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
'***
'***    HISTORY:
'***    Date        Description
'***    ----------  --------------------------------------------------------------------
'***    09/25/2003  Initial Creation
'***
'***************************************************************************************

Dim strOutput As String
Dim lngCounter As Long
Dim lngMaxStrLength As Long
lngMaxStrLength = Len(strStringList)

Randomize Timer

For lngCounter = 1 To lngLength
    
    strOutput = strOutput & Mid(strStringList, CLng((lngMaxStrLength - 2) * Rnd + 1), 1)
    
Next

    GetRandomString = strOutput

End Function


