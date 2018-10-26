Public Function hex2ascii(ByVal hextext As String) As String
    
For y = 1 To Len(hextext)
    num = Mid(hextext, y, 2)
    Value = Value & Chr(Val("&h" & num))
    y = y + 1
Next y

hex2ascii = Value
End Function