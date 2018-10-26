Public Function NumToString(ByVal nNumber As Currency) As String

Dim bNegative As Boolean
Dim bHundred As Boolean

If nNumber < 0 Then
    bNegative = True
End If

nNumber = Abs(Int(nNumber))

If nNumber < 1000 Then
    If nNumber \ 100 > 0 Then
        NumToString = NumToString & _
             NumToString(nNumber \ 100) & " hundred"
        bHundred = True
    End If
    nNumber = nNumber - ((nNumber \ 100) * 100)
    Dim bNoFirstDigit As Boolean
    bNoFirstDigit = False
    Select Case nNumber \ 10
        Case 0
            Select Case nNumber Mod 10
                Case 0
                    If Not bHundred Then
                        NumToString = NumToString & " zero"
                    End If
                Case 1: NumToString = NumToString & " one"
                Case 2: NumToString = NumToString & " two"
                Case 3: NumToString = NumToString & " three"
                Case 4: NumToString = NumToString & " four"
                Case 5: NumToString = NumToString & " five"
                Case 6: NumToString = NumToString & " six"
                Case 7: NumToString = NumToString & " seven"
                Case 8: NumToString = NumToString & " eight"
                Case 9: NumToString = NumToString & " nine"
            End Select
            bNoFirstDigit = True
        Case 1
            Select Case nNumber Mod 10
                Case 0: NumToString = NumToString & " ten"
                Case 1: NumToString = NumToString & " eleven"
                Case 2: NumToString = NumToString & " twelve"
                Case 3: NumToString = NumToString & " thirteen"
                Case 4: NumToString = NumToString & " fourteen"
                Case 5: NumToString = NumToString & " fifteen"
                Case 6: NumToString = NumToString & " sixteen"
                Case 7: NumToString = NumToString & " seventeen"
                Case 8: NumToString = NumToString & " eighteen"
                Case 9: NumToString = NumToString & " nineteen"
            End Select
            bNoFirstDigit = True
        Case 2: NumToString = NumToString & " twenty"
        Case 3: NumToString = NumToString & " thirty"
        Case 4: NumToString = NumToString & " forty"
        Case 5: NumToString = NumToString & " fifty"
        Case 6: NumToString = NumToString & " sixty"
        Case 7: NumToString = NumToString & " seventy"
        Case 8: NumToString = NumToString & " eighty"
        Case 9: NumToString = NumToString & " ninety"
    End Select
    If Not bNoFirstDigit Then
        If nNumber Mod 10 <> 0 Then
            NumToString = NumToString & "-" & _
                          Mid(NumToString(nNumber Mod 10), 2)
        End If
    End If
Else
    Dim nTemp As Currency
    nTemp = 10 ^ 12 'trillion
    Do While nTemp >= 1
        If nNumber >= nTemp Then
            NumToString = NumToString & _
                          NumToString(Int(nNumber / nTemp))
            Select Case Int(Log(nTemp) / Log(10) + 0.5)
                Case 12: NumToString = NumToString & " trillion"
                Case 9: NumToString = NumToString & " billion"
                Case 6: NumToString = NumToString & " million"
                Case 3: NumToString = NumToString & " thousand"
            End Select
           
            nNumber = nNumber - (Int(nNumber / nTemp) * nTemp)
        End If
        nTemp = nTemp / 1000
    Loop
End If

If bNegative Then
    NumToString = " negative" & NumToString
End If
    
End Function

Public Function DollarToString(ByVal nAmount As Currency) As _
String

    Dim nDollar As Currency
    Dim nCent As Currency
    
    nDollar = Int(nAmount)
    nCent = (Abs(nAmount) * 100) Mod 100
    
    DollarToString = NumToString(nDollar) & " dollar"
    
    If Abs(nDollar) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
    DollarToString = DollarToString & " and" & _
                     NumToString(nCent) & " cent"
                     
    If Abs(nCent) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
End Function
