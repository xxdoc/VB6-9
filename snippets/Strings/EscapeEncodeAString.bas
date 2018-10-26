Public Function Encode(ByVal Val As String) As String
    Dim i As Long, curChar As String

    For i = 1 To Len(Val)
        curChar = Mid(Val, i, 1)
        If curChar = vbCr Then
            Encode = Encode & "\c"
        ElseIf curChar = vbLf Then
            Encode = Encode & "\l"
        ElseIf curChar = "\" Then
            Encode = Encode & "\\"
        Else
            Encode = Encode & curChar
        End If
    Next i

End Function

Public Function Decode(ByVal Val As String) As String
    Dim i As Long, curChar As String, EscapeMode As Boolean
    
    For i = 1 To Len(Val)
        curChar = Mid(Val, i, 1)
        If EscapeMode = False Then
            If curChar = "\" Then
                EscapeMode = True
            Else
                Decode = Decode & curChar
            End If
        Else
            If curChar = "\" Then
                Decode = Decode & "\"
            ElseIf curChar = "c" Then
                Decode = Decode & vbCr
            ElseIf curChar = "l" Then
                Decode = Decode & vbLf
            End If
            EscapeMode = False
        End If
    Next i
    
End Function