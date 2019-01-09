VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmInventMaint 
   Caption         =   "Inventory Maintenance"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7320
      TabIndex        =   24
      Top             =   3120
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5160
      TabIndex        =   21
      Top             =   4200
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "On Sale?"
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Taxable"
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   5160
      TabIndex        =   12
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5160
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Qty. in stock "
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Discount %"
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Price"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Size"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Short Desc."
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Long Desc."
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Manufacturer"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "UPC"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmInventMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'define database data types
Dim rs As Recordset
Dim db As Database
Dim ws As Workspace

Dim MyEditRecord As String
Dim AddNew As Boolean

Private Sub Command1_Click()

AddNew = True

ClearFields

End Sub

Private Sub Command2_Click()
Dim Work As String
Dim Kount As Long
Dim ret As Long

ret = MsgBox("Are your sure you want to delete this record?")

If ret = 7 Then Exit Sub

rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1
    
    Work = rs("index")
    
    If Work = MyEditRecord Then
        rs.Delete
    End If
    
    rs.MoveNext
    
Next

LoadGrid

End Sub

Private Sub Command3_Click()
Dim Kount As Long
Dim Work As String
Dim ret As Long

If Text1.text = "" Then
    ret = MsgBox("You must enter a UPC (barcode) number!", vbOKOnly, "Error")
    Exit Sub
End If

If Combo1.text = "" Then Combo1.text = "Misc"

If Text3.text = "" Then Text3.text = "-"

If Text5.text = "" Then
    ret = MsgBox("You must enter a short description of the product!", vbOKOnly, "Error")
    Exit Sub
End If

If Text4.text = "" Then Text4.text = Text5.text

If Text6.text = "" Then Text6.text = "-"

If Text7.text = "" Then
    ret = MsgBox("You must enter a price for the product!", vbOKOnly, "Error")
    Exit Sub
End If

If Text2.text = "" Then Text2.text = "0"

If Text8.text = "" Then Text8.text = "0"

If AddNew = True Then
    
    rs.AddNew
        
    rs("barcode") = Text1.text
    rs("quantityinstock") = Text2.text
    rs("manufacturer") = Text3.text
    rs("longdesc") = Text4.text
    rs("shortdesc") = Text5.text
    rs("size") = Text6.text
    rs("price") = Text7.text
    rs("discountperc") = Text8.text
    rs("category") = Combo1.text
    
    If Check1.Value = vbChecked Then
        rs("taxable") = True
    Else
        rs("taxable") = False
    End If
    
    If Check2.Value = vbChecked Then
        rs("onsale") = True
    Else
        rs("onsale") = False
    End If
    
    rs.Update

Else
    
    rs.MoveFirst
    
    For Kount = 0 To rs.RecordCount - 1
        
        Work = rs("index")
        
        If Work = MyEditRecord Then
        
            rs.Edit
    
            rs("barcode") = Text1.text
            rs("quantityinstock") = Text2.text
            rs("manufacturer") = Text3.text
            rs("longdesc") = Text4.text
            rs("shortdesc") = Text5.text
            rs("size") = Text6.text
            rs("price") = Text7.text
            rs("discountperc") = Text8.text
            rs("category") = Combo1.text
            
            If Check1.Value = vbChecked Then
                rs("taxable") = True
            Else
                rs("taxable") = False
            End If
            
            If Check2.Value = vbChecked Then
                rs("onsale") = True
            Else
                rs("onsale") = False
            End If
        
            rs.Update
            
        End If
        
        rs.MoveNext
        
    Next
    
End If
    
DoEvents
AddNew = False
LoadGrid

End Sub

Private Sub Command5_Click()

Unload Me

End Sub

Private Sub Form_Load()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\main.mdb")
Set rs = db.OpenRecordset("tblmain", dbOpenTable)

LoadGrid
LoadCategories

End Sub

Private Sub LoadGrid()
Dim Kount As Long
Dim Work As String

Grid1.Clear
Grid1.Rows = 1

rs.MoveFirst

For Kount = 0 To rs.RecordCount - 1

    Work = rs("index") & Chr$(9) & rs("Barcode") & Chr$(9) & rs("category") & _
      Chr$(9) & rs("shortdesc") & Chr$(9) & rs("size") & Chr$(9) & _
      rs("price") & Chr$(9) & rs("quantityinstock")
      
    Grid1.AddItem Work
    
    rs.MoveNext
    
Next

'Grid1.RemoveItem (1)

Grid1.Col = 1
Grid1.Row = 0
Grid1.text = "UPC"

Grid1.Col = 2
Grid1.text = "Category"

Grid1.Col = 3
Grid1.text = "Description"

Grid1.Col = 4
Grid1.text = "Size"

Grid1.Col = 5
Grid1.text = "Price"

Grid1.Col = 6
Grid1.text = "Quantity"

Grid1.ColAlignment(1) = flexAlignCenterCenter
Grid1.ColAlignment(2) = flexAlignCenterCenter
Grid1.ColAlignment(3) = flexAlignLeftCenter
Grid1.ColAlignment(4) = flexAlignRightCenter
Grid1.ColAlignment(5) = flexAlignRightCenter
Grid1.ColAlignment(6) = flexAlignCenterCenter

End Sub

Private Sub Form_Resize()

On Error GoTo Hell

Command1.Left = 50
Command1.Top = Me.ScaleHeight - (Command1.Height + 50)

Command2.Left = Command1.Left + Command1.Width
Command2.Top = Command1.Top

Command3.Left = Command2.Left + Command2.Width
Command3.Top = Command2.Top

Command4.Left = Command3.Left + Command3.Width
Command4.Top = Command3.Top

Command5.Left = Command4.Left + Command4.Width
Command5.Top = Command4.Top

Label5.Top = Command1.Top - (Label1.Height + 50)
Label5.Left = Command1.Left

Label4.Top = Label5.Top - (Label4.Height + 50)
Label4.Left = Label5.Left

Label3.Top = Label4.Top - (Label3.Height + 50)
Label3.Left = Label4.Left

Label2.Top = Label3.Top - (Label2.Height + 50)
Label2.Left = Label3.Left

Label1.Top = Label2.Top - (Label1.Height + 50)
Label1.Left = Label2.Left

Text1.Left = Label1.Left + Label1.Width
Combo1.Left = Label2.Left + Label2.Width
Text3.Left = Label3.Left + Label3.Width
Text4.Left = Label4.Left + Label4.Width
Text5.Left = Label5.Left + Label5.Width

Text1.Top = Label1.Top
Combo1.Top = Label2.Top
Text3.Top = Label3.Top
Text4.Top = Label4.Top
Text5.Top = Label5.Top

Label8.Top = Label5.Top
Label8.Left = Text5.Left + Text5.Width + 500

Text8.Top = Label8.Top
Text8.Left = Label8.Left + Label8.Width

Check2.Left = Text8.Left
Check2.Top = Text8.Top - (Check2.Height + 50)

Check1.Left = Check2.Left
Check1.Top = Check2.Top - (Check1.Height + 50)

Label7.Left = Label8.Left
Label7.Top = Check1.Top - (Label7.Height + 50)

Text7.Left = Label7.Left + Label7.Width
Text7.Top = Label7.Top

Label6.Top = Label7.Top - (Label6.Height + 50)
Label6.Left = Label7.Left

Text6.Top = Label6.Top
Text6.Left = Label6.Left + Label6.Width

Label9.Top = Text8.Top
Label9.Left = Text8.Left + Text8.Width

Text2.Top = Label9.Top
Text2.Left = Label9.Left + Label9.Width

Grid1.Left = 0
Grid1.Width = Me.Width - 150
Grid1.Top = 0
 
Grid1.Height = Text1.Top - 50

Grid1.ColWidth(0) = 1
Grid1.ColWidth(1) = 1500
Grid1.ColWidth(2) = 1000
Grid1.ColWidth(3) = Grid1.Width - 5500
Grid1.ColWidth(4) = 1000
Grid1.ColWidth(5) = 750
Grid1.ColWidth(6) = 1000
Exit Sub

Hell:

End Sub

Private Sub Form_Unload(Cancel As Integer)

db.Close

frmMainMenu.Visible = True

End Sub

Private Sub LoadCategories()
Dim Kount As Long

Combo1.Clear
Combo1.text = ""

For Kount = 0 To NumCats
    Combo1.AddItem Categories(Kount)
Next


End Sub
Private Sub Grid1_Click()
Dim Work As String
Dim Kount As Long

ClearFields

Command2.Enabled = True

AddNew = False

Grid1.Col = 0
MyEditRecord = Grid1.text

rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1

    If rs("index") = MyEditRecord Then
    
        On Error Resume Next
        
        Text1.text = rs("barcode")
        Text2.text = rs("quantityinstock")
        Text3.text = rs("manufacturer")
        Text4.text = rs("longdesc")
        Text5.text = rs("shortdesc")
        Text6.text = rs("size")
        Text7.text = rs("price")
        Text8.text = rs("discountperc")
        Combo1.text = rs("category")
        
        On Error GoTo 0
        
        If rs("taxable") = True Then
            Check1.Value = vbChecked
        Else
            Check1.Value = vbUnchecked
        End If
        
        If rs("onsale") = True Then
            Check2.Value = vbChecked
        Else
            Check2.Value = vbUnchecked
        End If
        
    End If
    
    rs.MoveNext

Next

Grid1.ColSel = 6

End Sub

Private Sub ClearFields()

Text1.text = ""
Text2.text = ""
Text3.text = ""
Text4.text = ""
Text5.text = ""
Text6.text = ""
Text7.text = ""
Text8.text = ""
Check1.Value = vbUnchecked
Check2.Value = vbUnchecked

End Sub
