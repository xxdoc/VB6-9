Private Sub cmdTranslate_Click()
    Call translateNATO(Text1.Text)
End Sub

Private Sub translateNATO(strMsg)
    'strMsg = InputBox("Enter text:")
    
    strWords = Array("Alpha", "Bravo", "Charlie", "Delta", _
                    "Echo", "Foxtrot", "Golf", "Hotel", _
                    "India", "Juliet", "Kilo", "Lima", _
                    "Mike", "November", "Oscar", "Papa", _
                    "Quebec", "Romeo", "Sierra", "Tango", _
                    "Uniform", "Victor", "Whiskey", "Xray", _
                    "Yankee", "Zulu")
                
    If strMsg <> "" Then
        For i = 1 To Len(strMsg)
            If (Asc(LCase(Mid(strMsg, i, 1))) >= 97) And (Asc(LCase(Mid(strMsg, i, 1))) <= 122) Then
                    strOut = strOut & "-" & strWords(Asc(LCase(Mid(strMsg, i, 1))) - 97)
            Else
                If IsNumeric(Mid(strMsg, i, 1)) Then
                    strOut = strOut & "-" & Mid(strMsg, i, 1)
                Else
                    strOut = strOut & "-"
                End If
            End If
        Next
        MsgBox (strMsg & vbNewLine & "---------------" & vbNewLine & Mid(strOut, 2))
    End If
End Sub