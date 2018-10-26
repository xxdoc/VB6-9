Public Function VowelCount(ByVal InputString As String) As Long

Dim v(9) As String 'Declare an array  of 10 elements 0 to 9
Dim vcount As Integer 'This variable will contain number of vowels
Dim flag As Long
Dim strLen As Long
Dim i As Integer


v(0) = "a" 'First element of array is assigned small a
v(1) = "i"
v(2) = "o"
v(3) = "u"
v(4) = "e"
v(5) = "A" 'Sixth element is assigned Capital A
v(6) = "I"
v(7) = "O"
v(8) = "U"
v(9) = "E"
strLen = Len(InputString)

For flag = 1 To strLen 'It will get every letter of entered string and loop
'will terminate when all letters have been examined

    For i = 0 To 9 'Takes every elment of v(9) one by one
         'Check if current letter is a vowel
        If Mid(InputString, flag, 1) = v(i) Then              
              vcount = vcount + 1 ' If letter is equal to vowel
                                  'then increment vcount by 1
        End If
    Next i 'Consider next value of v(i)
Next flag 'Consider next letter of the enterd string
        
VowelCount = vcount

End Function