Public Function StrPalindrome(ByVal sSource As String) As String
    Dim l As Long, strRev As String
    Dim ab() As Byte
    ab = StrConv(sSource, vbFromUnicode)
    For l = UBound(ab) To 0 Step -1
        StrRev = StrRev & Chr$(ab(l))
    Next l
    If StrRev = sSource Then
        MsgBox "This is already a palindrome!"
    Else
        MsgBox "The palindrome of this word is" & vbCrLf & sSource & StrRev, vbOKOnly, "Palindrome Maker"
    End If
End Function