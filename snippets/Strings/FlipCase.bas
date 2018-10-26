Public Function FlipCase(ByVal ThisText As String) As String

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
        FlipCase= Temp
End Function
