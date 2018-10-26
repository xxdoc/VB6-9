Public Function StringStartsWith(ByVal strValue As String, _
  CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean
   
'Determines if a string starts with the same characters as 
'CheckFor string

'True if starts with CheckFor, false otherwise
'Case sensitive by default.  If you want non-case sensitive, set
'last parameter to vbTextCompare
    
    'Examples:
    'MsgBox StringStartsWith("Test", "TE") 'false
    'MsgBox StringStartsWith("Test", "TE", vbTextCompare) 'True
    
  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Left(strValue, lLen)
  StringStartsWith = StrComp(sCompare, CheckFor, CompareType) = 0

End Function

Public Function StringEndsWith(ByVal strValue As String, _
   CheckFor As String, Optional CompareType As VbCompareMethod _
   = vbBinaryCompare) As Boolean
 'Determines if a string ends with the same characters as
 'CheckFor string
 
 'True if end with CheckFor, false otherwise

 'Case sensitive by default.  If you want non-case sensitive, set
 'last parameter to vbTextCompare
 
  'Examples 
  'MsgBox StringEndsWith("Test", "ST") 'False
  'MsgBox StringEndsWith("Test", "ST", vbTextCompare) 'True

  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(CheckFor)
  If lLen > Len(strValue) Then Exit Function
  sCompare = Right(strValue, lLen)
  StringEndsWith = StrComp(sCompare, CheckFor, CompareType) = 0

End Function