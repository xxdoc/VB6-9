Private Function ByteArrayToString(Bytes() As Byte) As String

    Dim iUnicode As Long, i As Long, j As Long

    On Error Resume Next
    i = UBound(Bytes)

    If (i < 1) Then
        'ANSI, just convert to unicode and return
        ByteArrayToString = StrConv(Bytes, vbUnicode)
        Exit Function
    End If

    i = i + 1

    ' ***
    ' *** Examine the first two bytes
    ' *** 

    CopyMemory iUnicode, Bytes(0), 2

    If iUnicode = Bytes(0) Then 'Unicode
        
        ' *** 
        ' *** Account for terminating null
        ' *** 

        If (i Mod 2) Then i = i - 1
        
        ' *** 
        ' *** Set up a buffer to recieve the string
        ' *** 

        ByteArrayToString = String$(i / 2, 0)
        
        ' *** 
        ' *** Copy to string
        ' *** 

        CopyMemory ByVal StrPtr(ByteArrayToString), Bytes(0), i
        
    Else 'ANSI
    
            ByteArrayToString = StrConv(Bytes, vbUnicode)

    End If

End Function

Private Function StringToByteArray(strInput As String, Optional bReturnAsUnicode As Boolean = True, Optional bAddNullTerminator As Boolean = False) As Byte()

    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long

    If bReturnAsUnicode Then

        ' *** 
        ' *** Number of bytes
        ' *** 

        lLenB = LenB(strInput)
        
        ' *** 
        ' *** Resize buffer, do we want terminating null?
        ' *** 
        
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        
        ' *** 
        ' *** Copy characters from string to byte array
        ' *** 

        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB

    Else
    
        ' *** 
        ' *** Num of characters
        ' *** 

        lLenB = Len(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        
        lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)

    End If

    StringToByteArray = bytBuffer

End Function

Private Sub Command1_Click()

    Dim bAnsi() As Byte
    Dim bUni() As Byte
    Dim str As String
    Dim i As Long

    str = "Convert"
    bAnsi = StringToByteArray(str, False)
    bUni = StringToByteArray(str)

    For i = 0 To UBound(bAnsi)
        Debug.Print "=" & bAnsi(i)
    Next

    Debug.Print "========"

    For i = 0 To UBound(bUni)
        Debug.Print "=" & bUni(i)
    Next

    Debug.Print "ANSI= " & ByteArrayToString(bAnsi)
    Debug.Print "UNICODE= " & ByteArrayToString(bUni)
    
    ' *** 
    ' *** Using StrConv to convert a Unicode character array directly
    ' *** will cause the resultant string to have extra embedded nulls
    ' *** reason, StrConv does not know the difference between Unicode and ANSI

    Debug.print "Resull= " & StrConv(bUni,VbUnicode)

End Sub