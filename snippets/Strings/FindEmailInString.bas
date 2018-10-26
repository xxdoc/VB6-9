Public Function FindEMailAddressesInStr(ByVal StringWithEmails As String) As List(Of String)
    Dim emailList As New List(Of String)
    Dim RegExMatch As Text.RegularExpressions.MatchCollection = _
        System.Text.RegularExpressions.Regex.Matches(StringWithEmails, _
        "\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*")

    For i As Integer = 0 To RegExMatch.Count - 1
        If emailList.Contains(RegExMatch(i).Value) = False Then
            emailList.Add(RegExMatch(i).Value)
        End If
    Next

    Return emailList
End Function