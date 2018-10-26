Public Function InsertComma(ST As String) As String
 Dim st1, st2(100)
 Dim pos As Integer
 Dim len1 As Integer
 Dim flag As Boolean
 Dim dec As Variant
 pos = 0
 len1 = Len(Trim(ST))
   If len1 > 3 Then
        If InStr(ST, ".") > 0 Then
            flag = True
            dec = Right(ST, Len(ST) - InStr(ST, ".") + 1)
            ST = Left(ST, InStr(ST, ".") - 1)
        Else
            flag = False
        End If
        If Len(ST) = 4 Then
            InsertComma = Mid(ST, 1, 1) + "," + Mid(ST, 2, Len(ST))  '".00"
        ElseIf Len(ST) > 4 Then
            st1 = Left(ST, Len(ST) - 1)
            For i = Len(st1) To 1 Step -2
                st2(pos) = Right(st1, 2)
                If Len(st1) > 2 Then
                    st1 = Mid(st1, 1, Len(st1) - 2)
                Else
                    st2(pos) = st1
                End If
                pos = pos + 1
                'insertcomma = st2
            Next
            i = 0
            InsertComma = ""
            For i = pos - 1 To 0 Step -1
                If i <> 0 Then
                    InsertComma = InsertComma & st2(i) + ","
                Else
                    InsertComma = InsertComma & st2(i)
                End If
            Next
            If flag = True Then
                InsertComma = InsertComma + Right(ST, 1) + dec
            Else
                InsertComma = InsertComma + Right(ST, 1)
            End If
        End If
    Else
        InsertComma = Trim(ST)
    End If
End Function