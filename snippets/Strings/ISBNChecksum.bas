Sub VerifyISBN(ByVal strISBN)

    If CalculateISBNChecksum(strISBN) = Right(strISBN, 1) Then
        Debug.Print "ISBN Verified."
    Else
        Debug.Print "ISBN Not Verified. [" + CalculateISBNChecksum(strISBN) + "]"
    End If

End Sub

Function CalculateISBNChecksum(ByVal strISBN)
    
    Dim lngIterator
    Dim lngChecksum
    Dim lngChecksumEven
    lngIterator = CLng(0)
    lngChecksum = CLng(0)
    lngChecksumEven = CLng(0)
    strISBN = Replace(strISBN, "-", "")
    Select Case Len(strISBN)
    
        Case 10
    
            For lngIterator = 1 To Len(strISBN) - 1
                lngChecksum = lngChecksum + (lngIterator * Val(Mid(strISBN, lngIterator, 1)))
            Next
    
            If lngChecksum Mod 11 = 10 Then
                CalculateISBNChecksum = "X"
            Else
                CalculateISBNChecksum = Trim(CStr(lngChecksum Mod 11))
            End If
    
        Case 13
    
            For lngIterator = 1 To Len(strISBN) - 1
    
                If lngIterator Mod 2 = 1 Then
                    lngChecksumEven = lngChecksumEven + Val(Mid(strISBN, lngIterator, 1))
                Else
                    lngChecksum = lngChecksum + Val(Mid(strISBN, lngIterator, 1))
                End If
    
            Next
    
            lngChecksum = (lngChecksum * 3)
            lngChecksum = lngChecksum + lngChecksumEven
    
            If (10 - (lngChecksum Mod 10)) = 10 Then
                CalculateISBNChecksum = "0"
            Else
                CalculateISBNChecksum = Trim(CStr(10 - (lngChecksum Mod 10)))
            End If
    
        Case Else
    
    End Select

End Function