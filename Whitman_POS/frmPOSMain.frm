VERSION 5.00
Begin VB.Form frmPOSMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Point of Sale"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "End &Transaction"
      Height          =   375
      Left            =   5640
      TabIndex        =   29
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit POS"
      Height          =   375
      Left            =   8040
      TabIndex        =   28
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Accept Payment"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   24
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   20
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   3
      Left            =   8880
      TabIndex        =   17
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   8880
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&New Transaction"
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enter Item"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00008406&
      Caption         =   "Taxable?"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   3
      Left            =   6600
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   5
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   4
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   3
      Text            =   "02862226"
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Label Label13 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   9855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tendered:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   7920
      TabIndex        =   23
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5640
      TabIndex        =   22
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Due:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode:"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Transaction Receipt:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   10095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00008406&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   5520
      Top             =   1560
      Width           =   4575
   End
End
Attribute VB_Name = "frmPOSMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'define database data types
Dim ls As Recordset  ' for transaction logging
Dim rs As Recordset  ' for transaction detail
Dim db As Database
Dim ws As Workspace

'object variables
Dim GrandTotal As Currency
Dim SubTotal As Currency
Dim TaxTotal As Currency

Dim GUID As String
Private Sub cmdExit_Click()

frmMainMenu.Visible = True

Unload Me

End Sub

Private Sub Command1_Click()
Dim Kount As Long
Dim Work As String
Dim MyBarcode As String
Dim IsTaxable As Boolean
Dim Found As Boolean
Dim Padder As String
Dim rPrice As String
Dim TopLine As String

MyBarcode = Text2(0).text
Found = False

'search for item by barcode

Label1.Caption = "Searching..."
DoEvents

If Text2(0).text = "" Then
    
    Label13.Caption = "No Item Entered!"
    Exit Sub
    
End If



rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1
    
    Work = rs("barcode")
    
    If Work = MyBarcode Then
        
        'found it
        
        Found = True
        
        'update item text fields
        Text2(1).text = rs("longdesc")
        Text2(2).text = rs("shortdesc")
        Text2(3).text = Format(rs("price"), "$###,##0.00")
        
        'update topline
        Label13.Caption = Left(Text2(2).text, 35) & " (" & Text2(3).text & ")"
        
        'if it is taxable, calculate and display tax
        
        IsTaxable = rs("taxable")
        
        If IsTaxable = True Then
            Text2(4).text = Format(Text2(3).text * TaxRate, "$###,##0.00")
            Check1.Value = vbChecked
        Else
            Text2(4).text = "$0.00"
            Check1.Value = vbUnchecked
        End If
        
        'add to transaction detail log
        
        Set ls = db.OpenRecordset("tbltransdetail", dbOpenTable)
        
        On Error Resume Next
        
        ls.AddNew
        
        ls("guid") = GUID
        ls("barcode") = rs("barcode")
        ls("price") = rs("price")
        ls("category") = rs("category")
        ls("taxable") = rs("taxable")
        ls("manufacturer") = rs("manufacturer")
        ls("longdesc") = rs("longdesc")
        ls("shortdesc") = rs("shortdesc")
        ls("size") = rs("size")
        ls("onsale") = rs("onsale")
        ls("discoutperc") = rs("discountperc")
        
        ls.Update
        
        On Error GoTo 0
        
        'ensure the description can fit in the field
        
        If Len(rs("shortdesc")) < 25 Then
            
            Padder = rs("shortdesc") & Space$(25 - Len(rs("shortdesc")))
            
        End If
        
        If Len(rs("shortdesc")) > 25 Then
            
            Padder = Left(rs("shortdesc"), 25)
            
        End If
        
        'right align the price
        
        rPrice = Text2(3).text
        If Len(rPrice) < 10 Then
            rPrice = Space$(10 - Len(rPrice)) & rPrice
        End If
        
        Padder = Padder & rPrice
        
        'if the item is taxable, put an asterisk to the left
        
        If IsTaxable = True Then
            Padder = "*" & Padder
        Else
            Padder = " " & Padder
        End If
        
        'finally, update the receipt.
        
        Post Text1, " " & MyBarcode
        Post Text1, Padder
        
    End If
    
    rs.MoveNext

Next

If Found = True Then

    'update display
    
    SubTotal = SubTotal + Text2(3).text
    TaxTotal = TaxTotal + Text2(4).text
    GrandTotal = SubTotal + TaxTotal
    
    Text3(0).text = Format(SubTotal, "$###,##0.00")
    Text3(1).text = Format(TaxTotal, "$###,##0.00")
    Text3(2).text = Format(GrandTotal, "$###,##0.00")
    
    Label1.Caption = "Total = " & Format(GrandTotal, "$###,##0.00")
    Text2(0).text = ""
    
    
Else

    Label13.Caption = "Item Not Found!"
    Label1.Caption = "Total = " & Format(GrandTotal, "$###,##0.00")
    Text2(0).text = ""
    
End If

    
End Sub

Private Sub Command2_Click()

Label13.Caption = "Enter Amount Tendered"
Text3(3).text = ""
Text3(3).SetFocus
Command4.Default = True

End Sub

Private Sub Command3_Click()
Dim Kount As Long

'New Customer Button



'create guid
GUID = GetRandomString(15, "ABCDEFGHIJKLMNOPQRSTUVWXYZ01234567890")

Text1.text = ""

Post Text1, ReceiptHeader

GrandTotal = 0
SubTotal = 0
TaxTotal = 0

Label1.Caption = "Total = " & Format(GrandTotal, "$###,##0.00")
Label13.Caption = "Scan First Item"

Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command3.Enabled = False

Command1.Default = True

For Kount = 0 To 4
    Text3(Kount).text = ""
Next

For Kount = 0 To 4
    Text2(Kount).text = ""
Next

Text2(0).SetFocus

End Sub
Private Sub Command4_Click()
Dim Tendered As Currency
Dim ChangeDue As Currency
Dim AmountDue As Currency

Dim rSub As String
Dim rTax As String
Dim rTot As String
Dim rTen As String
Dim rCha As String

'calculate change due, and display it

If Text3(2).text = "" Or Text3(3).text = "" Then Exit Sub

AmountDue = Text3(2).text
Tendered = Text3(3).text

ChangeDue = Tendered - AmountDue

'format currencies
rSub = Format(SubTotal, "$###,##0.00")
rTax = Format(TaxTotal, "$###,##0.00")
rTot = Format(GrandTotal, "$###,##0.00")
rTen = Format(Tendered, "$###,##0.00")
rCha = Format(ChangeDue, "$###,##0.00")

'format displayed fields
Text3(3).text = rTen
Text3(4).text = rCha

'right align currencies for receipt
If Len(rSub) < 12 Then rSub = Space$(12 - Len(rSub)) & rSub
If Len(rTax) < 12 Then rTax = Space$(12 - Len(rTax)) & rTax
If Len(rTot) < 12 Then rTot = Space$(12 - Len(rTot)) & rTot
If Len(rTen) < 12 Then rTen = Space$(12 - Len(rTen)) & rTen
If Len(rCha) < 12 Then rCha = Space$(12 - Len(rCha)) & rCha

'update receipt, printing totals, amount tendered,
'change due, and the date/time.  footer is placed
'under the numbers. ("Have a Nice Day")

Post Text1, " "
Post Text1, "========================================"
Post Text1, "Subtotal = " & rSub
Post Text1, "Tax      = " & rTax
Post Text1, "Total    = " & rTot
Post Text1, " "
Post Text1, "Tendered = " & rTen
Post Text1, "Change   = " & rCha
Post Text1, " "
Post Text1, "========================================"
Post Text1, "          " & Date$ & "  " & Time$

If ShowGUID = True Then
    Post Text1, Space$(19 - (Len(GUID) / 2)) & GUID
End If

Post Text1, "========================================"

Post Text1, " "
If ShowFooter = True Then

    Post Text1, ReceiptFooter
    
End If

'update transaction summary table
ls.Close
DoEvents
Set ls = db.OpenRecordset("tblTransSummary", dbOpenTable)
DoEvents

On Error Resume Next
ls.AddNew

ls("GUID") = GUID
ls("subtotal") = SubTotal
ls("tax") = TaxTotal
ls("total") = GrandTotal
ls("tendered") = Tendered
ls("change") = ChangeDue
ls("date") = Date$
ls("time") = Time$

ls.Update
On Error GoTo 0
    

'update topline

Label1.Caption = "Change Due = " & Format(ChangeDue, "$###,##0.00")
Label13.Caption = "Amount Tendered = " & Format(Tendered, "$###,##0.00")

'change buttons

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = False
Command3.Default = True

End Sub

Private Sub Form_Load()

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\main.mdb")
Set rs = db.OpenRecordset("tblmain", dbOpenTable)


Label1.Caption = BigGreeting
Label13.Caption = "Ready for new transaction..."
Command3.Default = True

Command2.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

'close db before exitting

db.Close

End Sub

Private Sub Text3_Change(Index As Integer)

Select Case Index

    Case 3
        
        Label1.Caption = Text3(3).text
        
End Select

        
End Sub
