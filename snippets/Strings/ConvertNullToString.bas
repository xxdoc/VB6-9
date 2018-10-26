Public Function FixNull(vMayBeNull As Variant) As String
   On Error Resume Next
   FixNull = vbNullString & vMayBeNull
End Function