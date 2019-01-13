VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check7 
      Caption         =   "Show GUID"
      Height          =   255
      Left            =   5280
      TabIndex        =   31
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   4320
      TabIndex        =   29
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   27
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Footer"
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Internet URL"
      Height          =   255
      Left            =   5280
      TabIndex        =   24
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Email"
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Phone"
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Address"
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   720
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tagline"
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   2160
      TabIndex        =   11
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   2160
      TabIndex        =   10
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   2160
      TabIndex        =   9
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   2160
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   2160
      TabIndex        =   7
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   3960
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Re-Order Level: "
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   4920
      X2              =   7440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label9 
      Caption         =   "Display following on Receipts:"
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Footer:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet URL (WWW):"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Main Phone Number:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Your ""Tagline"":"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of your Business:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'define database data types
Dim rs As Recordset
Dim db As Database
Dim ws As Workspace

Private Sub Command1_Click()

Dim Kount As Long
Dim Work As String

'check for empty fields.  if empty, put a blank space in there

Work = " "

For Kount = 0 To 12

    If Text1(Kount).text = "" Then
        Text1(Kount).text = Work
    End If
    
Next
DoEvents

On Error Resume Next

'update settings in db

rs.MoveFirst
rs.Edit

rs("bizname") = Text1(0).text
rs("biztagline") = Text1(1).text
rs("address1") = Text1(2).text
rs("address2") = Text1(3).text
rs("city") = Text1(4).text
rs("State") = Text1(5).text
rs("zip") = Text1(6).text
rs("phone1") = Text1(7).text
rs("email") = Text1(8).text
rs("www") = Text1(9).text
rs("receiptfooter") = Text1(10).text
rs("taxrate") = Text1(11).text
rs("reorder") = Text1(12).text

If Check1.Value = vbUnchecked Then
    rs("showtagline") = False
Else
    rs("showtagline") = True
End If

If Check2.Value = vbUnchecked Then
    rs("showaddress") = False
Else
    rs("showaddress") = True
End If

If Check3.Value = vbUnchecked Then
    rs("showphone") = False
Else
    rs("showphone") = True
End If

If Check4.Value = vbUnchecked Then
    rs("showemail") = False
Else
    rs("showemail") = True
End If

If Check5.Value = vbUnchecked Then
    rs("showwww") = False
Else
    rs("showwww") = True
End If

If Check6.Value = vbUnchecked Then
    rs("showfooter") = False
Else
    rs("showfooter") = True
End If

If Check7.Value = vbUnchecked Then
    rs("showguid") = False
Else
    rs("showguid") = True
End If

rs.Update

On Error GoTo 0

're-load settings from db

rs.MoveFirst

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

ReOrderLevel = rs("reorder")

Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\main.mdb")
Set rs = db.OpenRecordset("tblsettings", dbOpenTable)


rs.MoveFirst
Text1(0).text = rs("bizname")
Text1(1).text = rs("biztagline")
Text1(2).text = rs("address1")
Text1(3).text = rs("address2")
Text1(4).text = rs("city")
Text1(5).text = rs("state")
Text1(6).text = rs("zip")
Text1(7).text = rs("phone1")
Text1(8).text = rs("email")
Text1(9).text = rs("www")
Text1(10).text = rs("receiptfooter")
Text1(11).text = rs("taxrate")
Text1(12).text = rs("reorder")

If rs("showtagline") = True Then
    Check1.Value = vbChecked
Else
    Check1.Value = vbUnchecked
End If

If rs("showaddress") = True Then
    Check2.Value = vbChecked
Else
    Check2.Value = vbUnchecked
End If

If rs("showphone") = True Then
    Check3.Value = vbChecked
Else
    Check3.Value = vbChecked
End If

If rs("showemail") = True Then
    Check4.Value = vbChecked
Else
    Check4.Value = vbUnchecked
End If

If rs("showwww") = True Then
    Check5.Value = vbChecked
Else
    Check5.Value = vbuncheckd
End If

If rs("showfooter") = True Then
    Check6.Value = vbChecked
Else
    Check6.Value = vbUnchecked
End If

If rs("showguid") = True Then
    Check7.Value = vbChecked
Else
    Check7.Value = vbUnchecked
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

db.Close

End Sub
