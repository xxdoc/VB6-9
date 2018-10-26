Public Function Encrypt(ByVal Plain As String, _
  sEncKey As String) As String
    '*********************************************************
    'Coded WhiteKnight 6-1-00
    'This Encrypts A string by converting it to its ASCII number
    'but the difference is it uses a Key String it converts the
    'keystring to ASCII and adds it to the first ASCII Value the
    'key is needed to decrypt the text.  I do plan on changing
    'this some what but For Now its ok.  I've only seen it
    'cause an error when the wrong Key was entered while
     'decrypting.
    
    'Note That If you use the same letter more then 3 times in a
    'row then each letter after it if still the same is ignored
    '(ie aaa = aaaaaaaaa but aaa <> aaaza)
    'If anyone Can figure out a way to fix this please e-mail me
  '*********************************************************
    Dim encrypted2 As String
    Dim LenLetter As Integer
    Dim Letter As String
    Dim KeyNum As String
    Dim encstr As String
    Dim temp As String
    Dim temp2 As String
    Dim itempstr As String
    Dim itempnum As Integer
    Dim Math As Long
    Dim i As Integer
    
    On Error GoTo oops
    
    If sEncKey = "" Then sEncKey = "WhiteKnight"
    'Sets the Encryption Key if one is not set
    ReDim encKEY(1 To Len(sEncKey))
    
    'starts the values for the Encryption Key
        
    For i = 1 To Len(sEncKey$)
     KeyNum = Mid$(sEncKey$, i, 1) 'gets the letter at index i
     encKEY(i) = Asc(KeyNum) 'sets the the Array value
                             'to ASC number for the letter

           'This is the first letter so just hold the value
        If i = 1 Then Math = encKEY(i): GoTo nextone

        'compares the value to the previous value and then
        'either adds/subtracts the value to the Math total
       If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) <= _
           encKEY(i - 1) Then Math = Math - encKEY(i)

        If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) <= _
           encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) >= Math And encKEY(i) >= _
           encKEY(i - 1) Then Math = Math + encKEY(i)
        If i >= 2 And encKEY(i) < Math And encKEY(i) _
          >= encKEY(i - 1) Then Math = Math + encKEY(i)
nextone:
    Next i
    
    
    For i = 1 To Len(Plain) 'Now for the String to be encrypted
        Letter = Mid$(Plain, i, 1) 'sets Letter to
                                   'the letter at index i
        LenLetter = Asc(Letter) + Math 'Now it adds the Asc
                                       'value of Letter to Math

'checks and corrects the format then adds a space to separate them frm each other
        If LenLetter >= 100 Then encstr = _
             encstr & Asc(Letter) + Math & " "

         'checks and corrects the format then adds a space
        'to separate them frm each other
        If LenLetter <= 99 Then encstr$ = encstr & "0" & _
          Asc(Letter) + Math & " "
    Next i


    'This is part of what i'm doing to convert the encrypted
    'numbers to Letters so it sort of encrypts the
    'encrypted message.
    temp$ = encstr 'hold the encrypted data
    temp$ = TrimSpaces(temp) 'get rid of the spaces
    itempnum% = Mid(temp, 1, 2) 'grab the first 2 numbers
    temp2$ = Chr(itempnum% + 100) 'Now add 100 so it
                                   'will be a valid char

    'If its a 2 digit number hold it and continue
    If Len(itempnum%) >= 2 Then itempstr$ = Str(itempnum%)
 
   'If the number is a single digit then add a '0' to the front
   'then hold it
    If Len(itempnum%) = 1 Then itempstr$ = "0" & _
        TrimSpaces(Str(itempnum%))
    
    encrypted2$ = temp2 'set the encrypted message
    
    For i = 3 To Len(temp) Step 2
        itempnum% = Mid(temp, i, 2) 'grab the next 2 numbers
  
      ' add 100 so it will be a valid char
        temp2$ = Chr(itempnum% + 100)

      'if its the last number we only want to hold it we
       'don't want to add a '0' even if its a single digit
        If i = Len(temp) Then itempstr$ = _
         Str(itempnum%): GoTo itsdone

'If its a 2 digit number hold it and continue
        If Len(itempnum%) = 2 Then itempstr$ = _
            Str(itempnum%)

        'If the number is a single digit then add a '0'
        'to the front then hold it
        If Len(TrimSpaces(Str(itempnum))) = 1 Then _
      itempstr$ = "0" & TrimSpaces(Str(itempnum%))

        'Now check to see if a - number was created
        'if so cause an error message
        If Left(TrimSpaces(Str(itempnum)), 1) = "-" Then _
          Err.Raise 20000, , "Unexpected Error"
           

itsdone:
           'Set The Encrypted message
        encrypted2$ = encrypted2 & temp2$
    Next i


    'Encrypt = encstr 'Returns the First Encrypted String
    Encrypt = encrypted2 'Returns the Second Encrypted String
    Exit Function 'We are outta Here
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
End Function

Public Function Decrypt(ByVal Encrypted As String, _
    sEncKey As String) As String

    Dim NewEncrypted As String
    Dim Letter As String
    Dim KeyNum As String
    Dim EncNum As String
    Dim encbuffer As Long
    Dim strDecrypted As String
    Dim Kdecrypt As String
    Dim lastTemp As String
    Dim LenTemp As Integer
    Dim temp As String
    Dim temp2 As String
    Dim itempstr As String
    Dim itempnum As Integer
    Dim Math As Long
    Dim i As Integer
    
    On Error GoTo oops

    If sEncKey = "" Then sEncKey = "WhiteKnight"

    ReDim encKEY(1 To Len(sEncKey))
    
    'Convert The Key For Decryption
    For i = 1 To Len(sEncKey$)
        KeyNum = Mid$(sEncKey$, i, 1) 'Get Letter i% in the Key
        encKEY(i) = Asc(KeyNum) 'Convert Letter i to Asc value
 
'if it the first letter just hold it
       If i = 1 Then Math = encKEY(i): GoTo nextone
       If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) _
               <= encKEY(i - 1) Then Math = Math - encKEY(i)
               'compares the value to the previous value and
               'then either adds/subtracts the value to the
               'Math total
        If i >= 2 And Math - encKEY(i) >= 0 And encKEY(i) _
              <= encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) >= Math And encKEY(i) _
              >= encKEY(i - 1) Then Math = Math + encKEY(i)
        If i >= 2 And encKEY(i) < Math And encKEY(i) _
              >= encKEY(i - 1) Then Math = Math + encKEY(i)
nextone:
    Next i
    
    
    'This is part of what i'm doing to convert the encrypted
    'numbers to  Letters so it sort of encrypts the encrypted
    'message.
    temp$ = Encrypted 'hold the encrypted data


    For i = 1 To Len(temp)
        itempstr = TrimSpaces(Str(Asc(Mid(temp, i, 1)) - _
           100)) 'grab the next 2 numbers
           'If its a 2 digit number hold it and continue
        If Len(itempstr$) = 2 Then itempstr$ = itempstr$
          If i = Len(temp) - 2 Then LenTemp% = _
               Len(Mid(temp2, Len(temp2) - 3))
          If i = Len(temp) Then itempstr$ = _
              TrimSpaces(itempstr$): GoTo itsdone
          'If the number is a single digit then add a '0' to the
          'front then hold it
        If Len(TrimSpaces(itempstr$)) = 1 Then _
             itempstr$ = "0" & TrimSpaces(itempstr$)
        'Now check to see if a - number was created if so
        'cause an error message
        If Left(TrimSpaces(itempstr$), 1) = "-" Then _
             Err.Raise 20000, , "Unexpected Error"
           

itsdone:
        temp2$ = temp2$ & itempstr 'hold the first decryption
    Next i
    
    
    Encrypted = TrimSpaces(temp2$) 'set the encrypted data


    For i = 1 To Len(Encrypted) Step 3
        'Format the encrypted string for the second decryption
        NewEncrypted = NewEncrypted & _
            Mid(Encrypted, CLng(i), 3) & " "
    Next i

' Hold the last set of numbers to check it its the correct format
    lastTemp$ = TrimSpaces(Mid(NewEncrypted, _
         Len(NewEncrypted$) - 3))
         
         If Len(lastTemp$) = 2 Then
' If it = 2 then its not the Correct format and we need to fix it
        lastTemp$ = Mid(NewEncrypted, _
           Len(NewEncrypted$) - 1) 'Holds Last Number so a '0'
                                    'Can be added between them
'set it to the new format
        Encrypted = Mid(NewEncrypted, 1, _
           Len(NewEncrypted) - 2) & "0" & lastTemp$
Else
        Encrypted$ = NewEncrypted$ 'set the new format

    End If
    'The Actual Decryption
    For i = 1 To Len(Encrypted)
        Letter = Mid$(Encrypted, i, 1) 'Hold Letter at index i
        EncNum = EncNum & Letter 'Hold the letters
        If Letter = " " Then 'we have a letter to decrypt
            encbuffer = CLng(Mid(EncNum, 1, _
              Len(EncNum) - 1)) 'Convert it to long and
                                 'get the number minus the " "
            strDecrypted$ = strDecrypted & Chr(encbuffer - _
               Math) 'Store the decrypted string
            EncNum = "" 'clear if it is a space so we can get
                        'the next set of numbers
        End If
    Next i

    Decrypt = strDecrypted

    Exit Function
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
Err.Raise 20001, , "You have entered the wrong encryption string"

End Function

Private Function TrimSpaces(strstring As String) As String
    Dim lngpos As Long
    Do While InStr(1&, strstring$, " ")
        DoEvents
         lngpos& = InStr(1&, strstring$, " ")
         strstring$ = Left$(strstring$, (lngpos& - 1&)) & _
            Right$(strstring$, Len(strstring$) - _
               (lngpos& + Len(" ") - 1&))
    Loop
     TrimSpaces$ = strstring$
End Function