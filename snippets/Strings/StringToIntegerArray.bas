' This array stores the array used as a lookup when translating 
' strings to arrays of integers.
Private mLookupArray(0 To 255) As Integer
Private Sub Class_Initialize()
    ' By default we'll use this array - it'll speed things up. 
    ' If you want another lookup you'll have to run the 
    ' GenerateRandomTranslation function. this is so it'll 
    ' always have some sort on conversion
    GenerateRandomTranslation
End Sub


Public Sub GenerateRandomTranslation(Optional ByVal theSeed _
                                      As Integer = 0)
    Dim ArrayIndex As Integer       ' The current array index
    Dim CheckIndex As Integer       ' The array index that 
                                    ' checks for repeats
    Dim UsedIndex As Integer        ' the used array index
    Dim UsedValues(0 To 255) As Boolean     ' The array of 
                                            ' used values
    
    ' Populate the array with random numbers. Don't worry if 
    ' the rounding off of the values means some of them repeat -
    ' we'll deal with that in a sec.
    For ArrayIndex = LBound(mLookupArray) To _
                                    UBound(mLookupArray)
        mLookupArray(ArrayIndex) = CInt(Rnd() * _
                                  CDbl(UBound(mLookupArray)))
    Next
        
    ' Reset the randomizer. This ensures we get the same random 
    ' numbers everytime we run it :)
    Rnd -1

    ' Randomize - This allows us to have up to 65355 lookup 
    ' tables - should be enough to start with
    Randomize theSeed
    
    ' Clear the array of used values. This should be done by 
    ' default - but it's a good thing to do (just in case)
    For ArrayIndex = LBound(UsedValues) To UBound(UsedValues)
        UsedValues(ArrayIndex) = False
    Next
    
    ' Now - we're going to define an array of values which are 
    ' used. This will then provide a list of values that aren't 
    ' used.
    For ArrayIndex = LBound(mLookupArray) To _
                                          UBound(mLookupArray)
        UsedValues(mLookupArray(ArrayIndex)) = True
    Next
    
    ' Now. We can step through the lookup array updating 
    ' repeated values with unused ones. Simple? Not really.
    For ArrayIndex = LBound(mLookupArray) To _
                                UBound(mLookupArray)
        For CheckIndex = LBound(mLookupArray) To _
                                    UBound(mLookupArray)
            ' If the values match - they are repeated. We must 
            ' add a new values that isn't used. A method is to 
            ' generate a new random number, and repeat 
            ' everything until we have a unquie list, trouble 
            ' is it's SLOW big time. Using this
            ' list of unused values is much quicker.
            If mLookupArray(ArrayIndex) = _
                             mLookupArray(CheckIndex) And _  
                             ArrayIndex <> CheckIndex Then
                ' Right - now find a free value - it works in 
                ' an order, but I am working
                ' on the assumption that the repeats are in all 
                ' sorts of places, not in an order.
                UsedIndex = 0
                Do While UsedValues(UsedIndex) = True
                    UsedIndex = UsedIndex + 1
                Loop
                ' Save it
                mLookupArray(ArrayIndex) = UsedIndex
                ' Mark it as used.
                UsedValues(UsedIndex) = True
            End If
        Next
    Next
    
End Sub


Public Sub Encode(ByVal theStringToEncode As String, _
                                ByRef theArray() As Integer)
    Dim IndexCount As Integer       ' String loop index

    ' Resize the array, so we have enough room to store all the 
    ' values. NOTE: this deletes anything that was there.
    ReDim theArray(Len(theStringToEncode) - 1)
    
    ' Now step through, anding the value of the ascii value 
    ' from the lookup table to the return array theArray.
    For IndexCount = 1 To Len(theStringToEncode)
        theArray(IndexCount - 1) = _
        mLookupArray(Asc(Mid(theStringToEncode, IndexCount, 1)))
    Next
End Sub


Public Function Decode(ByRef theArray() As Integer) As String
    Dim IndexCount As Integer   ' The current value to work on
    
    ' Clear the return value - not _really_ required, always a 
    ' good habit to get into.
    Decode = ""
    
    ' Loop round the array, determing the ascii value for each 
    ' entry and rebuild the string.
    For IndexCount = LBound(theArray) To UBound(theArray)
        Decode = Decode & _
               Chr(LocateValueFromLookup(theArray(IndexCount)))
    Next
End Function


Private Function LocateValueFromLookup(theValue As Integer) _
                                              As Integer
    Dim ArrayIndex As Integer       ' The index of the array
       
    ' Default return value - failure to find it shouldn't occur
    LocateValueFromLookup = 0
    
    ' Look at each entry in the lookup array and see if it 
    ' matches the value we're after. If it does, return the 
    ' index within the array for that value.
    For ArrayIndex = LBound(mLookupArray) To _
                                          UBound(mLookupArray)
        If mLookupArray(ArrayIndex) = theValue Then
            LocateValueFromLookup = ArrayIndex
        End If
    Next
End Function

Sample Usage:
Dim o As New clsEncoder
Dim iArr() As Integer, iCtr As Integer

Dim s As String
s = "This is a test"
o.Encode s, iArr
For iCtr = 0 To UBound(iArr)
    Debug.Print iArr(iCtr)
Next
MsgBox (o.Decode(iArr))