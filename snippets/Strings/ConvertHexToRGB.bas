Public Function HEXCOL2RGB(ByVal HexColor As String) As String

    'The input at this point could be HexColor = "#00FF1F"

Dim Red As String
Dim Green As String
Dim Blue As String

Color = Replace(HexColor, "#", "") 
    'Here HexColor = "00FF1F"

Red = Val("&H" & Mid(HexColor, 1, 2)) 
    'The red value is now the long version of "00"

Green = Val("&H" & Mid(HexColor, 3, 2))
    'The red value is now the long version of "FF"

Blue = Val("&H" & Mid(HexColor, 5, 2))
    'The red value is now the long version of "1F"


HEXCOL2RGB = RGB(Red, Green, Blue)
    'The output is an RGB value

End Function