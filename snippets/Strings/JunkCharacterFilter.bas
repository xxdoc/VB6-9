Public Function StrValidate(strSearch As Variant) As String

Dim rVal
rVal = strSearch
OPFst = JunkFilter(rVal)
OPFstSplit = Split(OPFst, "~*~", -1, 1)
OPFstVal = OPFstSplit(0)
SJunk = SJunk + OPFstSplit(1)

OPSec = FirstCheck(OPFstVal)
OPSecSplit = Split(OPSec, "~*~", -1, 1)
OPSecVal = OPSecSplit(0)
SJunk = SJunk + OPSecSplit(1)

OPThrd = LastCheck(OPSecVal)
OPThrdSplit = Split(OPThrd, "~*~", -1, 1)
OPThrdVal = OPThrdSplit(0)
SJunk = SJunk + OPThrdSplit(1)

'MsgBox (OPThrdVal)

OPFrth = StrDblCheck(OPThrdVal)
OPFrthSplit = Split(OPFrth, "~*~", -1, 1)
OPFrthVal = OPFrthSplit(0)
SJunk = SJunk + OPFrthSplit(1)

OPFifth = StrDblCheck1(OPFrthVal)
OPFifthSplit = Split(OPFifth, "~*~", -1, 1)
OPFifthVal = OPFifthSplit(0)
SJunk = SJunk + OPFifthSplit(1)

OPSxth = StrDblCheck(OPFifthVal)
OPSxthSplit = Split(OPSxth, "~*~", -1, 1)
OPSxthVal = OPSxthSplit(0)
SJunk = SJunk + OPSxthSplit(1)

StrValidate = OPSxthVal + "<A>" + SJunk

End Function


Private Function JunkFilter(iString As Variant) As Variant

For i = 1 To Len(iString)

StrRaw = Mid(iString, i, 1)
    
If StrRaw = "!" Or StrRaw = "@" Or StrRaw = "#" Or StrRaw = "$" Or StrRaw = "%" Or StrRaw = "^" Or StrRaw = "&" Or StrRaw = "*" Or StrRaw = "(" Or StrRaw = ")" Or StrRaw = "_" Or StrRaw = "-" Or StrRaw = "=" Or StrRaw = "|" Or StrRaw = "\" Or StrRaw = "[
    strJunk = strJunk + StrRaw
Else
    strKeyWords = strKeyWords + StrRaw
End If
Next i
JunkFilter = strKeyWords & "~*~" & strJunk

End Function

Private Function FirstCheck(strFirst As Variant) As String
    strFirstCheck = Mid(strFirst, 1, 1)
    While strFirstCheck = "+" Or strFirstCheck = "," Or strFirstCheck = " " Or strFirstCheck = "."
        strJunk = strJunk + strFirstCheck
        strFirst = Mid(strFirst, 2, Len(strFirst))
        strFirstCheck = Mid(strFirst, 1, 1)
    Wend
    FirstCheck = strFirst & "~*~" & strJunk
End Function

Private Function LastCheck(strLast As Variant) As String
    strLastCheck = Mid(strLast, Len(strLast), Len(strLast))
    While strLastCheck = "+" Or strLastCheck = " " Or strLastCheck = "," Or strLastCheck = "."
        strJunk = strJunk + strLastCheck
        strLast = Mid(strLast, 1, Len(strLast) - 1)
        strLastCheck = Mid(strLast, Len(strLast), Len(strLast))
    Wend
    LastCheck = strLast & "~*~" & strJunk
End Function

Private Function StrDblCheck(strDbl As Variant) As String

For k = 1 To Len(strDbl)

sStrDbl = Mid(strDbl, k, 1)
If k > 1 Then
    StrMinus = Mid(strDbl, k - 1, 1)
End If

If (sStrDbl = "+" And (StrMinus = "+" Or StrMinus = "," Or StrMinus = ".")) Or (sStrDbl = "," And (StrMinus = "+" Or StrMinus = "," Or StrMinus = ".")) Or (sStrDbl = "." And (StrMinus = "+" Or StrMinus = "," Or StrMinus = ".")) Or (sStrDbl = " " And (Str
    strJunk = strJunk + sStrDbl
Else
    strDblKey = strDblKey + sStrDbl
End If

Next

StrDblCheck = strDblKey & "~*~" & strJunk

End Function


Private Function StrDblCheck1(strDbl1 As Variant) As String

For k1 = 1 To Len(strDbl1)

sStrDbl1 = Mid(strDbl1, k1, 1)

strPlus = Mid(strDbl1, k1 + 1, 1)

If (sStrDbl1 = " " And (strPlus = "+" Or strPlus = "," Or strPlus = ".")) Then  ' Or StrMinus = "," Or StrMinus = ".")) Or (StrDblCheck = "," And (StrMinus = "+" Or StrMinus = "," Or StrMinus = ".")) Or (StrDblCheck = "." And (StrMinus = "+" Or StrMinus 
    strJunk = strJunk + sStrDbl1
Else
    strDblKey1 = strDblKey1 + sStrDbl1
End If

Next

StrDblCheck1 = strDblKey1 & "~*~" & strJunk

End Function