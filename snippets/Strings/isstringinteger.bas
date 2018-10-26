Public Function StringIsInteger(testString As String) As _
 Boolean

Dim asStr As String
Dim stuvv As String

'''' Note: The sections that strip off leading +, - , _
''''0 etc could be controlled by a second input parameter

stuvv = Trim(testString)
isAnIntegerType = False

'''''''' Leading + or -

If (InStr(stuvv, "+") = 1 Or InStr(stuvv, "-") = 1) _
 Then

    stuvv = Mid(stuvv, 2)
End If

'''''' This allow Hex -- remove if you don't want

If (InStr(stuvv, "0x") = 1 Or InStr(stuvv, "0X") = 1) _
   Then

    stuvv = Mid(stuvv, 3)

End If

'''''' This allows one leading zero( Octal )

If (Len(stuvv) > 1 And InStr(stuvv, "0") = 1) Then

    stuvv = Mid(stuvv, 2)

End If


'''''' This allows multiple leading zeros (leaves one
''''''if number is '0')


While (Len(stuvv) > 1 And InStr(stuvv, "0") = 1)

    stuvv = Mid(stuvv, 2)

Wend


If IsNumeric(stuvv) Then

    asStr = CStr(CInt(stuvv))

    If (stuvv = asStr) Then
        StringIsInteger = True
    End If


End If

End Function