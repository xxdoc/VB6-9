Attribute VB_Name = "modInitialize"
Global TaxRate As Double
Global ReceiptHeader As String
Global ReceiptFooter As String
Global StoreName As String
Global StorePhone As String
Global StoreWWW As String
Global StoreEmail As String
Global Tagline As String

Global Address As String
Global Address1 As String
Global address2 As String
Global City As String
Global State As String
Global ZIP As String

Global BigGreeting As String
Global Categories(5000) As String
Global NumCats As Long

Global ShowTagLine As Boolean
Global ShowAddress As Boolean
Global ShowPhone As Boolean
Global ShowEmail As Boolean
Global ShowWWW As Boolean
Global ShowFooter As Boolean
Global ShowGUID As Boolean

Global ReOrderLevel As Long
Sub Main()

'define database data types
Dim rs As Recordset
Dim db As Database
Dim ws As Workspace

Dim Kount As Long
Dim Work As String

Dim CurHour As Long

Dim Ctr As Long

'Greeting for Big Display
CurHour = Val(Left(Time, 2))

BigGreeting = "Good Evening"
If CurHour < 12 Then BigGreeting = "Good Morning"
If CurHour > 11 And CurHour < 17 Then BigGreeting = "Good Afternoon"

'set database/workspace
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\main.mdb")


'load settings
Set rs = db.OpenRecordset("tblsettings", dbOpenTable)
rs.MoveFirst

ReOrderLevel = rs("reorder")

ShowTagLine = rs("showtagline")
ShowAddress = rs("showaddress")
ShowPhone = rs("showphone")
ShowEmail = rs("showemail")
ShowWWW = rs("showwww")
ShowFooter = rs("showfooter")
ShowGUID = rs("showguid")

Tagline = rs("biztagline")

ReceiptFooter = rs("receiptfooter")

StoreName = rs("bizname")
Address1 = rs("address1")
address2 = rs("address2")

City = rs("city")
State = rs("state")
ZIP = rs("zip")

StorePhone = rs("phone1")
StoreWWW = rs("www")
StoreEmail = rs("email")

'this next section centers strings

Ctr = 19 - (Len(Tagline) / 2)
Tagline = Space$(Ctr) & Tagline

Ctr = 19 - (Len(ReceiptFooter) / 2)
ReceiptFooter = Space$(Ctr) & ReceiptFooter

Ctr = 19 - (Len(StoreName) / 2)
StoreName = Space$(Ctr) & StoreName

Ctr = 19 - (Len(Address1) / 2)
Address1 = Space$(Ctr) & Address1

Ctr = 19 - (Len(address2) / 2)
address2 = Space$(Ctr) & address2

Work = City & ", " & State & "  " & ZIP

Ctr = 19 - (Len(Work) / 2)
Work = Space$(Ctr) & Work

Ctr = 19 - (Len(StorePhone) / 2)
StorePhone = Space$(Ctr) & StorePhone

Ctr = 19 - (Len(StoreWWW) / 2)
StoreWWW = Space$(Ctr) & StoreWWW

Ctr = 19 - (Len(StoreEmail) / 2)
StoreEmail = Space$(Ctr) & StoreEmail

Address = Address1 & vbCrLf & address2 & vbCrLf & Work

'build reciept header

ReceiptHeader = StoreName & vbCrLf

If ShowAddress = True Then
    ReceiptHeader = ReceiptHeader & Address & vbCrLf
End If

If ShowPhone = True Then
    ReceiptHeader = ReceiptHeader & StorePhone & vbCrLf
End If

If ShowEmail = True Then
    ReceiptHeader = ReceiptHeader & StoreEmail & vbCrLf
End If

If ShowWWW = True Then
    ReceiptHeader = ReceiptHeader & StoreWWW & vbCrLf
End If

ReceiptHeader = ReceiptHeader & "========================================"

If ShowTagLine = True Then
    ReceiptHeader = ReceiptHeader & vbCrLf & Tagline & _
       vbCrLf & "========================================"
End If

'if the user doesn't want to show the footer, change it to null

If ShowFooter = False Then
    ReceiptFooter = ""
End If

TaxRate = rs("taxrate")
TaxRate = TaxRate / 100

'load categories
Set rs = db.OpenRecordset("tblcategories", dbOpenTable)
rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1

    Categories(Kount) = rs("category")
    
    rs.MoveNext

Next
NumCats = rs.RecordCount - 1
db.Close

'show main menu
frmMainMenu.Visible = True
frmSplash.Visible = True

End Sub
