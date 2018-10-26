Function FormatByLength(Expression As String, Length As Long) _
   As String

   Dim BufferCrLf() As String
   Dim BufferSpace() As String
   Dim Buffer As String
   Dim k As Long
   Dim j As Long
   Dim count As Long
On Error GoTo FormatByLengthError
   BufferCrLf() = Split(Expression, vbCrLf)
   For k = 0 To UBound(BufferCrLf())
       If Len(BufferCrLf(k)) <= Length Then
          Buffer = Buffer & BufferCrLf(k) & vbCrLf
       Else
          BufferSpace() = Split(BufferCrLf(k), " ")
          For j = 0 To UBound(BufferSpace())
              count = count + Len(BufferSpace(j)) + 1
              If (count <= Length) Then
                 Buffer = Buffer & BufferSpace(j) & " "
              Else
                 count = 0
                 Buffer = Buffer & vbCrLf & BufferSpace(j) & " "
                 count = Len(BufferSpace(j)) + 1
              End If
          Next j
          Buffer = Buffer & vbCrLf
       End If
   Next k
   FormatByLength = Buffer
   Exit Function
FormatByLengthError:
    
End Function