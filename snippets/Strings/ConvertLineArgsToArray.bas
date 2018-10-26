'---------------------------------------------------------------------------------------
' CommandLineToArray - Alex Jephson 2008
'---------------------------------------------------------------------------------------
' This function is used to split a set of input parameters into an array.
' VB6 Code Tested Only
'---------------------------------------------------------------------------------------
' Inputs
' ------
' strCommand:    A string of characters which is the input parameter of the program.
' blnSaveQuotes: A boolean which saves the quotes from the input string.
' (Optional)     If this parameter is omitted false is assumed.
'
' If a zero-length string is passed to the function a zero length string will be
' returned in element 0 of the array.
'---------------------------------------------------------------------------------------
' Outputs
' -------
' Function returns an Variant (array of variable length strings); 1 element for each
' parameter (0 Minimum Bound)
'---------------------------------------------------------------------------------------
' Example
' -------
' arrRet = CommandLineToArray("""This is my"" parameter string", False)
' arrRet(0): This is my
' arrRet(1): parameter
' arrRet(2): string
' --or--
' arrRet = CommandLineToArray("""This is my"" parameter string", True)
' arrRet(0): "This is my"
' arrRet(1): parameter
' arrRet(2): string
'---------------------------------------------------------------------------------------

Function CommandLineToArray(strCommand As String, Optional blnSaveQuotes As Boolean) As Variant
    Dim arrParam() As String
    Dim blnInQuotes As Boolean
    Dim lngLength As String
    Dim iCount, iLoop As Integer
    Dim strCurrentChar As String
    
    lngLength = Len(strCommand)
    iCount = 0
    ReDim Preserve arrParam(iCount)
    
    For iLoop = 1 To lngLength
        strCurrentChar = Mid$(strCommand, iLoop, 1)
        If strCurrentChar = """" Then
            If blnInQuotes Then
                blnInQuotes = False
            Else
                blnInQuotes = True
            End If
            If blnSaveQuotes Then arrParam(iCount) = arrParam(iCount) + strCurrentChar
        ElseIf strCurrentChar = " " And Not blnInQuotes Then
            iCount = iCount + 1
            ReDim Preserve arrParam(iCount)
        Else
            arrParam(iCount) = arrParam(iCount) + strCurrentChar
        End If
    Next iLoop
    
    CommandLineToArray = arrParam
End Function