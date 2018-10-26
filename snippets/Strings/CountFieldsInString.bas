Function CountFields(LineDataIN As String, _
     Delimiter As String) As Integer
    Dim NewPos As Integer
    Dim MaxPos As Integer
    Dim FieldCounter As Integer
    
    If LineDataIN = "" Or Delimiter = "" Then
        CountFields = 0
        Exit Function
    End If
    
    MaxPos = Len(LineDataIN)
    NewPos = 1
    FieldCounter = 1
    
    While (NewPos < MaxPos) And (NewPos <> 0)
        NewPos = InStr(NewPos, LineDataIN, _
            Delimiter, vbTextCompare)
        If NewPos <> 0 Then
            FieldCounter = FieldCounter + 1
            NewPos = NewPos + 1
        End If
    Wend
    CountFields = FieldCounter
End Function


Sub CmdTest_Click
'This is just a test routine and isn't required....
'Just here to show how the code can be used.

     Dim NumberOfFields As Integer
     Dim Delimiter As String
     Dim LineDataIn As String
     
     Delimiter = Inputbox$("Type a Field Delimiter","DEMO:Delimited Field Counter",",")
     LineDataIn = Inputbox$("Enter a Delimted String to have its fields counted","DEMO:Delimited Field Counter")
     
     NumberOfFields = CountFields(LineDataIN, Delimiter)
     Msgbox "There are: " + Str$(NumberOfFields) + " fields in the string ("+LineDataIn + ")",64,"DEMO:Delimited Field Counter"
End Sub