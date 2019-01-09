VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PersonBuilder"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPerson 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Person"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblPersonData 
      Caption         =   "Person Data:"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curLocationCount As Currency        ' number of locations (currency for size)
Dim curStreetTypeCount As Currency      ' number of street types (currency for size)
Dim curStreetCount As Currency          ' number of street names (currency for size)
Dim curFemaleNameCount As Currency      ' number of female names (currency for size)
Dim curMaleNameCount As Currency        ' number of male names (currency for size)
Dim curSurnameCount As Currency         ' number of surnames (currency for size)

Dim strLocations()                      ' (tab delimited) location data including zip code, city name, state, county, latitude, longitude, and timezone.
Dim strSurnames()                       ' (single field) list of surnames
Dim strMaleNames()                      ' (single field) list of male names
Dim strFemaleNames()                    ' (single field) list of female names
Dim strStreets()                        ' (single field) list of street names
Dim strStreetTypes()                    ' (single field) list of street types (street, blvd, etc)

Private Sub LoadStreetTypes()

' **************************************************************************************
' **
' ** Load Street Types
' **
' **************************************************************************************

    Dim strWork As String                   ' temporary holding string
    Dim curKount As Currency                ' temporary big number holding
       
    ' *** First, get a count of the locations in the file.
    
    Open App.Path & "\PB-StreetTypes.txt" For Input As #1
    
    While Not EOF(1)
        Line Input #1, strWork
        curStreetTypeCount = curStreetTypeCount + 1
    Wend
    
    Close #1
    
    ' *** ReDimension an array with the number of locations
    
    ReDim strStreetTypes(curStreetTypeCount)
    
    ' *** Load the array
    
    Open App.Path & "\PB-StreetTypes.txt" For Input As #1
    curKount = 0
    
    While Not EOF(1)
    
        Line Input #1, strWork
        strStreetTypes(curKount) = strWork
        curKount = curKount + 1

    Wend
    
    Close #1

End Sub
Private Sub LoadStreets()

' **************************************************************************************
' **
' ** Load Streets
' **
' **************************************************************************************

    Dim strWork As String                   ' temporary holding string
    Dim curKount As Currency                ' temporary big number holding
       
    ' *** First, get a count of the locations in the file.
    
    Open App.Path & "\PB-Streets.txt" For Input As #1
    
    While Not EOF(1)
        Line Input #1, strWork
        curStreetCount = curStreetCount + 1
    Wend
    
    Close #1
    
    ' *** ReDimension an array with the number of locations
    
    ReDim strStreets(curStreetCount)
    
    ' *** Load the array
    
    Open App.Path & "\PB-Streets.txt" For Input As #1
    curKount = 0
    
    While Not EOF(1)
        Line Input #1, strWork
        strStreets(curKount) = strWork
        curKount = curKount + 1

    Wend
    
    Close #1
    
End Sub
Private Sub LoadFemaleNames()

' **************************************************************************************
' **
' ** Load female names
' **
' **************************************************************************************

    Dim strWork As String                   ' temporary holding string
    Dim curKount As Currency                ' temporary big number holding
    
    ' *** First, get a count of the locations in the file.
    
    Open App.Path & "\PB-Female.txt" For Input As #1
    
    While Not EOF(1)
        Line Input #1, strWork
        curFemaleNameCount = curFemaleNameCount + 1
    Wend
    
    Close #1
    
    ' *** ReDimension an array with the number of locations
    
    ReDim strFemaleNames(curFemaleNameCount)
    
    ' *** Load the array
    
    Open App.Path & "\PB-Female.txt" For Input As #1
    curKount = 0
    
    While Not EOF(1)
        Line Input #1, strWork
        
        ' ensure proper case before adding to array
        strWork = UCase(Left(strWork, 1)) & Right(strWork, Len(strWork) - 1)
        strFemaleNames(curKount) = strWork
        curKount = curKount + 1

    Wend
    
    Close #1

End Sub
Private Sub LoadMaleNames()

' **************************************************************************************
' **
' ** Load male names
' **
' **************************************************************************************

    Dim strWork As String                   ' temporary holding string
    Dim curKount As Currency                ' temporary big number holding
    
    ' *** First, get a count of the locations in the file.
    
    Open App.Path & "\PB-Male.txt" For Input As #1
    
    While Not EOF(1)
        Line Input #1, strWork
        curMaleNameCount = curMaleNameCount + 1
    Wend
    
    Close #1
    
    ' *** ReDimension an array with the number of locations
    
    ReDim strMaleNames(curMaleNameCount)
    
    ' *** Load the array
    
    Open App.Path & "\PB-Male.txt" For Input As #1
    curKount = 0
    
    While Not EOF(1)
        Line Input #1, strWork
        ' ensure proper case before adding to array
        strWork = UCase(Left(strWork, 1)) & Right(strWork, Len(strWork) - 1)
        strMaleNames(curKount) = strWork
        curKount = curKount + 1

    Wend
    
    Close #1
    
End Sub
Private Sub LoadSurnames()
    
' **************************************************************************************
' **
' ** Load Surnames
' **
' **************************************************************************************
    
    Dim strWork As String                   ' temporary holding string
    Dim curKount As Currency                ' temporary big number holding
    
    ' *** First, get a count of the locations in the file.
    
    Open App.Path & "\PB-Surnames.txt" For Input As #1
    
    While Not EOF(1)
        Line Input #1, strWork
        curSurnameCount = curSurnameCount + 1
    Wend
    
    Close #1
    
    ' *** ReDimension an array with the number of locations
    
    ReDim strSurnames(curSurnameCount)
    
    ' *** Load the array
    
    Open App.Path & "\PB-Surnames.txt" For Input As #1
    curKount = 0
    
    While Not EOF(1)
        Line Input #1, strWork
        ' ensure proper case before adding to array
        strWork = UCase(Left(strWork, 1)) & Right(strWork, Len(strWork) - 1)
        strSurnames(curKount) = strWork
        curKount = curKount + 1

    Wend
    
    Close #1
    
End Sub
Private Sub LoadLocations()

' **************************************************************************************
' **
' ** Load Locations
' **
' **************************************************************************************

    Dim strWork As String                   ' temporary holding string
    Dim curKount As Currency                ' temporary big number holding
    
        
    ' *** First, get a count of the locations in the file.
    
    Open App.Path & "\PB-Locations.txt" For Input As #1
    
    While Not EOF(1)
        Line Input #1, strWork
        curLocationCount = curLocationCount + 1
    Wend
    
    Close #1
    
    ' *** ReDimension an array with the number of locations
    
    ReDim strLocations(curLocationCount)
    
    ' *** Load the array
    
    Open App.Path & "\PB-Locations.txt" For Input As #1
    curKount = 0
    
    While Not EOF(1)
        Line Input #1, strWork
        strLocations(curKount) = strWork
        curKount = curKount + 1

    Wend
    
    Close #1
        
End Sub

Private Sub cmdGenerate_Click()

' **************************************************************************************
' **
' ** Generate a person
' **
' **************************************************************************************

    Dim MyRandom As Currency                ' Using currency for this HUGE random number
    Dim Kount As Long                       ' for loops
    Dim strNewPerson As String              ' The final data for the new person
    Dim strMyLocation As String             ' Location
    Dim strMyGender As String               ' for picking the gender (male/female)
    Dim strMyGivenName As String            ' First name, chosen from male or female list, based on intMyGender value
    Dim strMyStreet As String               ' street name
    Dim strMyStreetType As String           ' ave, blvd, street, etc.
    Dim strMySurname As String              ' Surname or family name
    Dim intMyHouseNumber As Long            ' house number
    Dim strMyPhoneNumber As String          ' randomized phone number
    Dim strAreaCode As String               ' area code (from Location)
    Dim strCity As String                   ' city (from Location)
    Dim strState As String                  ' state (from Location)
    Dim strZIP As String                    ' zip code (from Location)

    ' *** To get this show on the road, we need to seed the random number generator. If
    ' *** this isn't done, we'll get the same person data generated on each run.
    
    Randomize Timer

    ' *** Pick a gender
    
    MyRandom = Int(Rnd(1) * 100) + 1
    If MyRandom < 50 Then strMyGender = "Male" Else strMyGender = "Female"
    
    ' *** Pick a given (first) name, based on gender
    Select Case strMyGender
        
        Case "Male"
            MyRandom = Int(Rnd(1) * UBound(strMaleNames))
            strMyGivenName = strMaleNames(MyRandom)
            
        Case "Female"
            MyRandom = Int(Rnd(1) * UBound(strFemaleNames))
            strMyGivenName = strFemaleNames(MyRandom)
        
    End Select
    
    ' *** Pick a surname
    
    MyRandom = Int(Rnd(1) * UBound(strSurnames))
    strMySurname = strSurnames(MyRandom)
    
    ' *** Pick a location
    
    MyRandom = Int(Rnd(1) * UBound(strLocations))
    strMyLocation = strLocations(MyRandom)
    
    ' *** Pick a street and type, and generate a house number
    
    MyRandom = Int(Rnd(1) * UBound(strStreets))
    strMyStreet = strStreets(MyRandom)
    
    MyRandom = Int(Rnd(1) * UBound(strStreetTypes))
    strMyStreetType = strStreetTypes(MyRandom)
    
    intMyHouseNumber = Int(Rnd(1) * 19999) + 100
    
    ' *** Each record in the location file actually contains eight (8) pieces of data:
    ' ***
    ' ***   - ZIP Code
    ' ***   - City name
    ' ***   - State
    ' ***   - County
    ' ***   - Area code
    ' ***   - Latitude
    ' ***   - Longitude
    ' ***   - Time zone
    ' ***
    ' *** This file contains 42,009 records, but 336,072 pieces of information. If you
    ' *** are thinking about putting this in a database, there are a few ways of
    ' *** optimizing for space utilization.
    ' ***
    ' *** For this program, all we need are the area code, city, state, and zip from
    ' *** the location. We'll use ParameterValue to extract the fields from the
    ' *** tab-delimited record.
    
    strAreaCode = Left(ParameterValue(Chr(9), strMyLocation, 5), 3)
    strCity = ParameterValue(Chr(9), strMyLocation, 2)
    strState = ParameterValue(Chr(9), strMyLocation, 3)
    strZIP = ParameterValue(Chr(9), strMyLocation, 1)
    
    ' *** generate a random phone number, starting with area code. just need to add 7 more digits
    
    strMyPhoneNumber = strAreaCode
    For Kount = 1 To 7
        MyRandom = Int(Rnd(1) * 9) + 1
        strMyPhoneNumber = strMyPhoneNumber & Trim(Str(MyRandom))
    Next
    
    strMyPhoneNumber = PhoneFormat(strMyPhoneNumber)
    
    ' *** String it all together, and plug it into the textbox
    
    strNewPerson = "Name:     " & strMyGivenName & " " & strMySurname & vbCrLf
    strNewPerson = strNewPerson & "Gender:   " & strMyGender & vbCrLf
    strNewPerson = strNewPerson & "Phone:    " & strMyPhoneNumber & vbCrLf
    strNewPerson = strNewPerson & vbCrLf
    strNewPerson = strNewPerson & "Address:  " & intMyHouseNumber & " " & strMyStreet & " " & strMyStreetType & vbCrLf
    strNewPerson = strNewPerson & "          " & strCity & ", " & strState & " " & strZIP & vbCrLf
    
    txtPerson.Text = strNewPerson
    
End Sub

Private Sub Form_Load()

' **************************************************************************************
' **
' ** Let's get loaded!
' **
' **************************************************************************************
    
    ' *** first, show the wait dialog box, then run the load routines.
    
    frmWait.Visible = True
    frmWait.lblTitle = "Please Wait"
    frmWait.lblDetail = "Loading Data!"
    frmWait.Caption = "Hold on!"
    
    DoEvents
    
    LoadLocations
    LoadSurnames
    LoadMaleNames
    LoadFemaleNames
    LoadStreets
    LoadStreetTypes

    ' *** aaaaand disappear the wait dialog
    
    Unload frmWait
    
End Sub
