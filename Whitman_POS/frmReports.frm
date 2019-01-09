VERSION 5.00
Begin VB.Form frmReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4455
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command9 
         Caption         =   "Get Details"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Select a transaction from the list, or type it into the box below:"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stock Reports"
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Re-Order Report"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "All Products"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Departments Products"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Sale Items"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Taxable Items"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Non-Taxable Items"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Select Dept."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Exit to Main Menu"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Manufacturers Products"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'define database data types
Dim rs As Recordset
Dim db As Database
Dim ws As Workspace
Dim ls As Recordset

Private Sub Command1_Click()

Dim Kount As Long
Dim Work As String
Dim MyQuant As Long

Dim rBarcode As String
Dim rCategory As String
Dim rManufacturer As String
Dim rShortDesc As String
Dim rSize As String
Dim rPrice As String
Dim rQty As String

Dim NumberOfItems As Long

NumberOfItems = 0

frmRptOutput.Visible = True
frmRptOutput.Text1.text = ""

'create header for report

frmRptOutput.Text1.text = "Re-order Report" & vbCrLf & Date$ & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "  UPC/Barcode #  Department  Manufacturer     Description      Size            Price   Qty" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " --------------  ----------  ---------------  ---------------  -------------  ------  ----" & vbCrLf

rs.MoveFirst

For Kount = 0 To rs.RecordCount - 1

    MyQuant = rs("quantityinstock")
    If MyQuant < ReOrderLevel Then
        
        On Error Resume Next
        
        rBarcode = ""
        rCategory = ""
        rManufacturer = ""
        rShortDesc = ""
        rSize = ""
        rPrice = ""
        rQty = ""
        
        rBarcode = rs("barcode")
        rCategory = rs("category")
        rManufacturer = rs("manufacturer")
        rShortDesc = rs("shortdesc")
        rSize = rs("size")
        rPrice = rs("price")
        rQty = rs("quantityinstock")
        
        On Error GoTo 0
        
        If Len(rBarcode) < 15 Then
            
            rBarcode = Space$(15 - Len(rBarcode)) & rBarcode
            
        End If
        
        If Len(rCategory) < 10 Then
            rCategory = rCategory & Space$(10 - Len(rCategory))
        End If
        
        If Len(rCategory) > 10 Then
            rCategory = Left(rCategory, 10)
        End If
        
        If Len(rManufacturer) < 15 Then
            rManufacturer = rManufacturer & Space$(15 - Len(rManufacturer))
        End If
        
        If Len(rManufacturer) > 15 Then
            rManufacturer = Left(manufacturer, 15)
        End If
        
        If Len(rShortDesc) < 15 Then
            rShortDesc = rShortDesc & Space$(15 - Len(rShortDesc))
        End If
        
        If Len(rShortDesc) > 15 Then
            rShortDesc = Left(rShortDesc, 15)
        End If
        
        If Len(rSize) < 8 Then
            rSize = rSize & Space$(8 - Len(rSize))
        End If
        
        If Len(rSize) > 8 Then
            rSize = Left(rSize, 8)
        End If
        
        rPrice = Format$(rPrice, "$###,##0.00")
        
        If Len(rPrice) < 11 Then
            rPrice = Space$(11 - Len(rPrice)) & rPrice
        End If
        
        rQty = Trim(rQty)
        If Len(rQty) < 4 Then
            rQty = Space$(4 - Len(rQty)) & rQty
        End If
        
        Work = rBarcode & "  " & rCategory & "  " & rManufacturer & "  " & _
           rShortDesc & "  " & rSize & "  " & rPrice & "  " & rQty & vbCrLf
           
        frmRptOutput.Text1.text = frmRptOutput.Text1.text & Work
        
        NumberOfItems = NumberOfItems + 1
        
    End If
    
    rs.MoveNext
    
Next

frmRptOutput.Text1.text = frmRptOutput.Text1.text & " -----------------------------------------------------------------------------------------" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " " & NumberOfItems & " items." & vbCrLf

End Sub
Private Sub Command2_Click()
Dim Kount As Long
Dim Work As String
Dim MyQuant As Long

Dim rBarcode As String
Dim rCategory As String
Dim rManufacturer As String
Dim rShortDesc As String
Dim rSize As String
Dim rPrice As String
Dim rQty As String

Dim NumberOfItems As Long

NumberOfItems = 0

frmRptOutput.Visible = True
frmRptOutput.Text1.text = ""

'create header for report

frmRptOutput.Text1.text = "All Products Report" & vbCrLf & Date$ & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "  UPC/Barcode #  Department  Manufacturer     Description      Size            Price   Qty" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " --------------  ----------  ---------------  ---------------  -------------  ------  ----" & vbCrLf

rs.MoveFirst

For Kount = 0 To rs.RecordCount - 1

    'MyQuant = rs("quantityinstock")
    'If MyQuant < ReOrderLevel Then
        
        rBarcode = " "
        rCategory = " "
        rManufacturer = " "
        rShortDesc = " "
        rSize = " "
        rPrice = " "
        rQty = " "
        
        On Error Resume Next
        rBarcode = rs("barcode")
        rCategory = rs("category")
        rManufacturer = rs("manufacturer")
        rShortDesc = rs("shortdesc")
        rSize = rs("size")
        rPrice = rs("price")
        rQty = rs("quantityinstock")
        
        On Error GoTo 0
        
        If Len(rBarcode) < 15 Then
            
            rBarcode = Space$(15 - Len(rBarcode)) & rBarcode
            
        End If
        
        If Len(rCategory) < 10 Then
            rCategory = rCategory & Space$(10 - Len(rCategory))
        End If
        
        If Len(rCategory) > 10 Then
            rCategory = Left(rCategory, 10)
        End If
        
        If Len(rManufacturer) < 15 Then
            rManufacturer = rManufacturer & Space$(15 - Len(rManufacturer))
        End If
        
        If Len(rManufacturer) > 15 Then
            rManufacturer = Left(manufacturer, 15)
        End If
        
        If Len(rShortDesc) < 15 Then
            rShortDesc = rShortDesc & Space$(15 - Len(rShortDesc))
        End If
        
        If Len(rShortDesc) > 15 Then
            rShortDesc = Left(rShortDesc, 15)
        End If
        
        If Len(rSize) < 8 Then
            rSize = rSize & Space$(8 - Len(rSize))
        End If
        
        If Len(rSize) > 8 Then
            rSize = Left(rSize, 8)
        End If
        
        rPrice = Format$(rPrice, "$###,##0.00")
        
        If Len(rPrice) < 11 Then
            rPrice = Space$(11 - Len(rPrice)) & rPrice
        End If
        
        rQty = Trim(rQty)
        If Len(rQty) < 4 Then
            rQty = Space$(4 - Len(rQty)) & rQty
        End If
        
        Work = rBarcode & "  " & rCategory & "  " & rManufacturer & "  " & _
           rShortDesc & "  " & rSize & "  " & rPrice & "  " & rQty & vbCrLf
           
        frmRptOutput.Text1.text = frmRptOutput.Text1.text & Work
        
        NumberOfItems = NumberOfItems + 1
        
    'End If
    
    rs.MoveNext
    
Next

frmRptOutput.Text1.text = frmRptOutput.Text1.text & " -----------------------------------------------------------------------------------------" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " " & NumberOfItems & " items." & vbCrLf

End Sub

Private Sub Command3_Click()
Dim Kount As Long
Dim Work As String
Dim MyManuf As String

Dim rBarcode As String
Dim rCategory As String
Dim rManufacturer As String
Dim rShortDesc As String
Dim rSize As String
Dim rPrice As String
Dim rQty As String

Dim NumberOfItems As Long

If Text1.text = "" Then Exit Sub

NumberOfItems = 0

frmRptOutput.Visible = True
frmRptOutput.Text1.text = ""

'create header for report

frmRptOutput.Text1.text = "Manufacturers Items Report" & vbCrLf & Date$ & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "  UPC/Barcode #  Department  Manufacturer     Description      Size            Price   Qty" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " --------------  ----------  ---------------  ---------------  -------------  ------  ----" & vbCrLf

rs.MoveFirst

For Kount = 0 To rs.RecordCount - 1

    On Error Resume Next
    MyManuf = rs("manufacturer")
    On Error GoTo 0
    
    If UCase(MyManuf) = UCase(Text1.text) And Len(MyManuf) > 1 Then
        
        On Error Resume Next
        
        rBarcode = ""
        rCategory = ""
        rManufacturer = ""
        rShortDesc = ""
        rSize = ""
        rPrice = ""
        rQty = ""
        
        rBarcode = rs("barcode")
        rCategory = rs("category")
        rManufacturer = rs("manufacturer")
        rShortDesc = rs("shortdesc")
        rSize = rs("size")
        rPrice = rs("price")
        rQty = rs("quantityinstock")
        
        On Error GoTo 0
        
        If Len(rBarcode) < 15 Then
            
            rBarcode = Space$(15 - Len(rBarcode)) & rBarcode
            
        End If
        
        If Len(rCategory) < 10 Then
            rCategory = rCategory & Space$(10 - Len(rCategory))
        End If
        
        If Len(rCategory) > 10 Then
            rCategory = Left(rCategory, 10)
        End If
        
        If Len(rManufacturer) < 15 Then
            rManufacturer = rManufacturer & Space$(15 - Len(rManufacturer))
        End If
        
        If Len(rManufacturer) > 15 Then
            rManufacturer = Left(manufacturer, 15)
        End If
        
        If Len(rShortDesc) < 15 Then
            rShortDesc = rShortDesc & Space$(15 - Len(rShortDesc))
        End If
        
        If Len(rShortDesc) > 15 Then
            rShortDesc = Left(rShortDesc, 15)
        End If
        
        If Len(rSize) < 8 Then
            rSize = rSize & Space$(8 - Len(rSize))
        End If
        
        If Len(rSize) > 8 Then
            rSize = Left(rSize, 8)
        End If
        
        rPrice = Format$(rPrice, "$###,##0.00")
        
        If Len(rPrice) < 11 Then
            rPrice = Space$(11 - Len(rPrice)) & rPrice
        End If
        
        rQty = Trim(rQty)
        If Len(rQty) < 4 Then
            rQty = Space$(4 - Len(rQty)) & rQty
        End If
        
        Work = rBarcode & "  " & rCategory & "  " & rManufacturer & "  " & _
           rShortDesc & "  " & rSize & "  " & rPrice & "  " & rQty & vbCrLf
           
        frmRptOutput.Text1.text = frmRptOutput.Text1.text & Work
        
        NumberOfItems = NumberOfItems + 1
        
    End If
    
    rs.MoveNext
    
Next

frmRptOutput.Text1.text = frmRptOutput.Text1.text & " -----------------------------------------------------------------------------------------" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " " & NumberOfItems & " items." & vbCrLf




End Sub

Private Sub Command4_Click()
Dim Kount As Long
Dim Work As String
Dim MyDept As String

Dim rBarcode As String
Dim rCategory As String
Dim rManufacturer As String
Dim rShortDesc As String
Dim rSize As String
Dim rPrice As String
Dim rQty As String

Dim NumberOfItems As Long

If Combo1.text = "" Then Exit Sub

NumberOfItems = 0

frmRptOutput.Visible = True
frmRptOutput.Text1.text = ""
Work = ""

'create header for report

frmRptOutput.Text1.text = "Departmental Report" & vbCrLf & Date$ & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "  UPC/Barcode #  Department  Manufacturer     Description      Size            Price   Qty" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " --------------  ----------  ---------------  ---------------  -------------  ------  ----" & vbCrLf

rs.MoveFirst

For Kount = 0 To rs.RecordCount - 1

    MyDept = rs("category")
    If MyDept = Combo1.text Then
        
        On Error Resume Next
        
        rBarcode = ""
        rCategory = ""
        rManufacturer = ""
        rShortDesc = ""
        rSize = ""
        rPrice = ""
        rQty = ""
        
        rBarcode = rs("barcode")
        rCategory = rs("category")
        rManufacturer = rs("manufacturer")
        rShortDesc = rs("shortdesc")
        rSize = rs("size")
        rPrice = rs("price")
        rQty = rs("quantityinstock")
        
        On Error GoTo 0
        
        If Len(rBarcode) < 15 Then
            
            rBarcode = Space$(15 - Len(rBarcode)) & rBarcode
            
        End If
        
        If Len(rCategory) < 10 Then
            rCategory = rCategory & Space$(10 - Len(rCategory))
        End If
        
        If Len(rCategory) > 10 Then
            rCategory = Left(rCategory, 10)
        End If
        
        If Len(rManufacturer) < 15 Then
            rManufacturer = rManufacturer & Space$(15 - Len(rManufacturer))
        End If
        
        If Len(rManufacturer) > 15 Then
            rManufacturer = Left(manufacturer, 15)
        End If
        
        If Len(rShortDesc) < 15 Then
            rShortDesc = rShortDesc & Space$(15 - Len(rShortDesc))
        End If
        
        If Len(rShortDesc) > 15 Then
            rShortDesc = Left(rShortDesc, 15)
        End If
        
        If Len(rSize) < 8 Then
            rSize = rSize & Space$(8 - Len(rSize))
        End If
        
        If Len(rSize) > 8 Then
            rSize = Left(rSize, 8)
        End If
        
        rPrice = Format$(rPrice, "$###,##0.00")
        
        If Len(rPrice) < 11 Then
            rPrice = Space$(11 - Len(rPrice)) & rPrice
        End If
        
        rQty = Trim(rQty)
        If Len(rQty) < 4 Then
            rQty = Space$(4 - Len(rQty)) & rQty
        End If
        
        Work = rBarcode & "  " & rCategory & "  " & rManufacturer & "  " & _
           rShortDesc & "  " & rSize & "  " & rPrice & "  " & rQty & vbCrLf
           
        frmRptOutput.Text1.text = frmRptOutput.Text1.text & Work
        
        NumberOfItems = NumberOfItems + 1
        
    End If
    
    rs.MoveNext
    
Next

frmRptOutput.Text1.text = frmRptOutput.Text1.text & " -----------------------------------------------------------------------------------------" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " " & NumberOfItems & " items." & vbCrLf
End Sub

Private Sub Command5_Click()
Dim Kount As Long
Dim Work As String
Dim OnSale As Boolean

Dim rBarcode As String
Dim rCategory As String
Dim rManufacturer As String
Dim rShortDesc As String
Dim rSize As String
Dim rPrice As String
Dim rQty As String
Dim rDiscount As Long
Dim MyDiscount As String

Dim NumberOfItems As Long

NumberOfItems = 0

frmRptOutput.Visible = True
frmRptOutput.Text1.text = ""
Work = ""

'create header for report

frmRptOutput.Text1.text = "Items On Sale" & vbCrLf & Date$ & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "                                                                                            Disc" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "  UPC/Barcode #  Department  Manufacturer     Description      Size            Price   Qty    %" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " --------------  ----------  ---------------  ---------------  -------------  ------  ----  ----" & vbCrLf

rs.MoveFirst

For Kount = 0 To rs.RecordCount - 1

    OnSale = rs("onsale")
    If OnSale = True Then
        
        On Error Resume Next
        
        rBarcode = ""
        rCategory = ""
        rManufacturer = ""
        rShortDesc = ""
        rSize = ""
        rPrice = ""
        rQty = ""
        rDiscount = 0
        
        rBarcode = rs("barcode")
        rCategory = rs("category")
        rManufacturer = rs("manufacturer")
        rShortDesc = rs("shortdesc")
        rSize = rs("size")
        rPrice = rs("price")
        rQty = rs("quantityinstock")
        rDiscount = rs("discountperc")
        
        On Error GoTo 0
        
        If Len(rBarcode) < 15 Then
            
            rBarcode = Space$(15 - Len(rBarcode)) & rBarcode
            
        End If
        
        If Len(rCategory) < 10 Then
            rCategory = rCategory & Space$(10 - Len(rCategory))
        End If
        
        If Len(rCategory) > 10 Then
            rCategory = Left(rCategory, 10)
        End If
        
        If Len(rManufacturer) < 15 Then
            rManufacturer = rManufacturer & Space$(15 - Len(rManufacturer))
        End If
        
        If Len(rManufacturer) > 15 Then
            rManufacturer = Left(manufacturer, 15)
        End If
        
        If Len(rShortDesc) < 15 Then
            rShortDesc = rShortDesc & Space$(15 - Len(rShortDesc))
        End If
        
        If Len(rShortDesc) > 15 Then
            rShortDesc = Left(rShortDesc, 15)
        End If
        
        If Len(rSize) < 8 Then
            rSize = rSize & Space$(8 - Len(rSize))
        End If
        
        If Len(rSize) > 8 Then
            rSize = Left(rSize, 8)
        End If
        
        rPrice = Format$(rPrice, "$###,##0.00")
        
        If Len(rPrice) < 11 Then
            rPrice = Space$(11 - Len(rPrice)) & rPrice
        End If
        
        rQty = Trim(rQty)
        If Len(rQty) < 4 Then
            rQty = Space$(4 - Len(rQty)) & rQty
        End If
        
        MyDiscount = Trim(Str(rDiscount / 100))
        MyDiscount = Format(MyDiscount, "##%")
                
        If Len(MyDiscount) < 4 Then
            MyDiscount = Space$(4 - Len(MyDiscount)) & MyDiscount
        End If
        
        Work = rBarcode & "  " & rCategory & "  " & rManufacturer & "  " & _
           rShortDesc & "  " & rSize & "  " & rPrice & "  " & rQty & "  " & MyDiscount & vbCrLf
           
        frmRptOutput.Text1.text = frmRptOutput.Text1.text & Work
        
        NumberOfItems = NumberOfItems + 1
        
    End If
    
    rs.MoveNext
    
Next

frmRptOutput.Text1.text = frmRptOutput.Text1.text & " -----------------------------------------------------------------------------------------------" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " " & NumberOfItems & " items." & vbCrLf

End Sub

Private Sub Command6_Click()
Dim Kount As Long
Dim Work As String
Dim MyDept As String

Dim rBarcode As String
Dim rCategory As String
Dim rManufacturer As String
Dim rShortDesc As String
Dim rSize As String
Dim rPrice As String
Dim rQty As String

Dim IsTaxable As Boolean

Dim NumberOfItems As Long

NumberOfItems = 0

frmRptOutput.Visible = True
frmRptOutput.Text1.text = ""
Work = ""

'create header for report

frmRptOutput.Text1.text = "Taxable Items Report" & vbCrLf & Date$ & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "  UPC/Barcode #  Department  Manufacturer     Description      Size            Price   Qty" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " --------------  ----------  ---------------  ---------------  -------------  ------  ----" & vbCrLf

rs.MoveFirst

For Kount = 0 To rs.RecordCount - 1

    IsTaxable = rs("taxable")
    If IsTaxable = True Then
        
        On Error Resume Next
        
        rBarcode = ""
        rCategory = ""
        rManufacturer = ""
        rShortDesc = ""
        rSize = ""
        rPrice = ""
        rQty = ""
        
        rBarcode = rs("barcode")
        rCategory = rs("category")
        rManufacturer = rs("manufacturer")
        rShortDesc = rs("shortdesc")
        rSize = rs("size")
        rPrice = rs("price")
        rQty = rs("quantityinstock")
        
        On Error GoTo 0
        
        If Len(rBarcode) < 15 Then
            
            rBarcode = Space$(15 - Len(rBarcode)) & rBarcode
            
        End If
        
        If Len(rCategory) < 10 Then
            rCategory = rCategory & Space$(10 - Len(rCategory))
        End If
        
        If Len(rCategory) > 10 Then
            rCategory = Left(rCategory, 10)
        End If
        
        If Len(rManufacturer) < 15 Then
            rManufacturer = rManufacturer & Space$(15 - Len(rManufacturer))
        End If
        
        If Len(rManufacturer) > 15 Then
            rManufacturer = Left(manufacturer, 15)
        End If
        
        If Len(rShortDesc) < 15 Then
            rShortDesc = rShortDesc & Space$(15 - Len(rShortDesc))
        End If
        
        If Len(rShortDesc) > 15 Then
            rShortDesc = Left(rShortDesc, 15)
        End If
        
        If Len(rSize) < 8 Then
            rSize = rSize & Space$(8 - Len(rSize))
        End If
        
        If Len(rSize) > 8 Then
            rSize = Left(rSize, 8)
        End If
        
        rPrice = Format$(rPrice, "$###,##0.00")
        
        If Len(rPrice) < 11 Then
            rPrice = Space$(11 - Len(rPrice)) & rPrice
        End If
        
        rQty = Trim(rQty)
        If Len(rQty) < 4 Then
            rQty = Space$(4 - Len(rQty)) & rQty
        End If
        
        Work = rBarcode & "  " & rCategory & "  " & rManufacturer & "  " & _
           rShortDesc & "  " & rSize & "  " & rPrice & "  " & rQty & vbCrLf
           
        frmRptOutput.Text1.text = frmRptOutput.Text1.text & Work
        
        NumberOfItems = NumberOfItems + 1
        
    End If
    
    rs.MoveNext
    
Next

frmRptOutput.Text1.text = frmRptOutput.Text1.text & " -----------------------------------------------------------------------------------------" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " " & NumberOfItems & " items." & vbCrLf

End Sub

Private Sub Command7_Click()
Dim Kount As Long
Dim Work As String
Dim MyDept As String

Dim rBarcode As String
Dim rCategory As String
Dim rManufacturer As String
Dim rShortDesc As String
Dim rSize As String
Dim rPrice As String
Dim rQty As String

Dim IsTaxable As Boolean

Dim NumberOfItems As Long

NumberOfItems = 0

frmRptOutput.Visible = True
frmRptOutput.Text1.text = ""
Work = ""

'create header for report

frmRptOutput.Text1.text = "Non-Taxable Items Report" & vbCrLf & Date$ & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "  UPC/Barcode #  Department  Manufacturer     Description      Size            Price   Qty" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " --------------  ----------  ---------------  ---------------  -------------  ------  ----" & vbCrLf

rs.MoveFirst

For Kount = 0 To rs.RecordCount - 1

    IsTaxable = rs("taxable")
    If IsTaxable = False Then
        
        On Error Resume Next
        
        rBarcode = ""
        rCategory = ""
        rManufacturer = ""
        rShortDesc = ""
        rSize = ""
        rPrice = ""
        rQty = ""
        
        rBarcode = rs("barcode")
        rCategory = rs("category")
        rManufacturer = rs("manufacturer")
        rShortDesc = rs("shortdesc")
        rSize = rs("size")
        rPrice = rs("price")
        rQty = rs("quantityinstock")
        
        On Error GoTo 0
        
        If Len(rBarcode) < 15 Then
            
            rBarcode = Space$(15 - Len(rBarcode)) & rBarcode
            
        End If
        
        If Len(rCategory) < 10 Then
            rCategory = rCategory & Space$(10 - Len(rCategory))
        End If
        
        If Len(rCategory) > 10 Then
            rCategory = Left(rCategory, 10)
        End If
        
        If Len(rManufacturer) < 15 Then
            rManufacturer = rManufacturer & Space$(15 - Len(rManufacturer))
        End If
        
        If Len(rManufacturer) > 15 Then
            rManufacturer = Left(manufacturer, 15)
        End If
        
        If Len(rShortDesc) < 15 Then
            rShortDesc = rShortDesc & Space$(15 - Len(rShortDesc))
        End If
        
        If Len(rShortDesc) > 15 Then
            rShortDesc = Left(rShortDesc, 15)
        End If
        
        If Len(rSize) < 8 Then
            rSize = rSize & Space$(8 - Len(rSize))
        End If
        
        If Len(rSize) > 8 Then
            rSize = Left(rSize, 8)
        End If
        
        rPrice = Format$(rPrice, "$###,##0.00")
        
        If Len(rPrice) < 11 Then
            rPrice = Space$(11 - Len(rPrice)) & rPrice
        End If
        
        rQty = Trim(rQty)
        If Len(rQty) < 4 Then
            rQty = Space$(4 - Len(rQty)) & rQty
        End If
        
        Work = rBarcode & "  " & rCategory & "  " & rManufacturer & "  " & _
           rShortDesc & "  " & rSize & "  " & rPrice & "  " & rQty & vbCrLf
           
        frmRptOutput.Text1.text = frmRptOutput.Text1.text & Work
        
        NumberOfItems = NumberOfItems + 1
        
    End If
    
    rs.MoveNext
    
Next

frmRptOutput.Text1.text = frmRptOutput.Text1.text & " -----------------------------------------------------------------------------------------" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " " & NumberOfItems & " items." & vbCrLf

End Sub

Private Sub Command8_Click()

frmMainMenu.Visible = True
Unload Me

End Sub

Private Sub Command9_Click()
Dim Kount As Long
Dim Work As String

Dim TransDate As String
Dim TransTime As String
Dim GUID As String
Dim SubTotal As String
Dim Tax As String
Dim GrandTotal As String
Dim Tendered As String
Dim ChangeDue As String

Dim rBarcode As String
Dim rCategory As String
Dim rManufacturer As String
Dim rShortDesc As String
Dim rSize As String
Dim rPrice As String
Dim rQty As String

Dim IsTaxable As Boolean

Dim NumberOfItems As Long

NumberOfItems = 0

frmRptOutput.Visible = True
frmRptOutput.Text1.text = ""
Work = ""

'create header for report

frmRptOutput.Text1.text = "Transaction Detail Report" & vbCrLf & Date$ & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & "  UPC/Barcode #  Department  Manufacturer     Description      Size            Price" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " --------------  ----------  ---------------  ---------------  -------------  ------" & vbCrLf


'first, get the info from the summary table

Set ls = db.OpenRecordset("tblTransSummary", dbOpenTable)

ls.MoveFirst

For Kount = 0 To ls.RecordCount - 1
    If ls("guid") = UCase(Text2.text) Then
        
        TransDate = ls("date")
        TransTime = ls("time")
        GUID = ls("guid")
        GUID = UCase(GUID)
        SubTotal = ls("subtotal")
        Tax = ls("tax")
        GrandTotal = ls("total")
        Tendered = ls("tendered")
        ChangeDue = ls("change")
        
    End If
    
    ls.MoveNext

Next

ls.Close

DoEvents

'now, get the info from the details table

Set ls = db.OpenRecordset("tblTransDetail", dbOpenTable)
        
ls.MoveFirst

For Kount = 0 To ls.RecordCount - 1
    
    If UCase(ls("guid")) = GUID Then
        
        On Error Resume Next
        
        rBarcode = ""
        rCategory = ""
        rManufacturer = ""
        rShortDesc = ""
        rSize = ""
        rPrice = ""
        rQty = ""
        
        rBarcode = ls("barcode")
        rCategory = ls("category")
        rManufacturer = ls("manufacturer")
        rShortDesc = ls("shortdesc")
        rSize = ls("size")
        rPrice = ls("price")
        
        On Error GoTo 0
        
        If Len(rBarcode) < 15 Then
            
            rBarcode = Space$(15 - Len(rBarcode)) & rBarcode
            
        End If
        
        If Len(rCategory) < 10 Then
            rCategory = rCategory & Space$(10 - Len(rCategory))
        End If
        
        If Len(rCategory) > 10 Then
            rCategory = Left(rCategory, 10)
        End If
        
        If Len(rManufacturer) < 15 Then
            rManufacturer = rManufacturer & Space$(15 - Len(rManufacturer))
        End If
        
        If Len(rManufacturer) > 15 Then
            rManufacturer = Left(manufacturer, 15)
        End If
        
        If Len(rShortDesc) < 15 Then
            rShortDesc = rShortDesc & Space$(15 - Len(rShortDesc))
        End If
        
        If Len(rShortDesc) > 15 Then
            rShortDesc = Left(rShortDesc, 15)
        End If
        
        If Len(rSize) < 8 Then
            rSize = rSize & Space$(8 - Len(rSize))
        End If
        
        If Len(rSize) > 8 Then
            rSize = Left(rSize, 8)
        End If
        
        rPrice = Format$(rPrice, "$###,##0.00")
        
        If Len(rPrice) < 11 Then
            rPrice = Space$(11 - Len(rPrice)) & rPrice
        End If
        
        rQty = Trim(rQty)
        If Len(rQty) < 4 Then
            rQty = Space$(4 - Len(rQty)) & rQty
        End If
        
        Work = rBarcode & "  " & rCategory & "  " & rManufacturer & "  " & _
           rShortDesc & "  " & rSize & "  " & rPrice & "  " & rQty & vbCrLf
           
        frmRptOutput.Text1.text = frmRptOutput.Text1.text & Work
        
        NumberOfItems = NumberOfItems + 1
        
    End If
    
    ls.MoveNext
    
Next

frmRptOutput.Text1.text = frmRptOutput.Text1.text & " -----------------------------------------------------------------------------------------" & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " " & NumberOfItems & " items." & vbCrLf & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " Subtotal:     " & Format(SubTotal, "$###,##0.00") & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " Tax     :     " & Format(Tax, "$###,##0.00") & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " Total   :     " & Format(GrandTotal, "$###,##0.00") & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " Tendered:     " & Format(Tendered, "$###,##0.00") & vbCrLf
frmRptOutput.Text1.text = frmRptOutput.Text1.text & " Change  :     " & Format(ChangeDue, "$###,##0.00") & vbCrLf


        

End Sub

Private Sub Form_Load()
Dim Kount As Long
Dim Work As String

For Kount = 0 To NumCats
    Combo1.AddItem Categories(Kount)
Next

List1.Clear

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\main.mdb")
Set rs = db.OpenRecordset("tblmain", dbOpenTable)
Set ls = db.OpenRecordset("tblTransSummary", dbOpenTable)

'load transaction GUIDS from the transaction summary table
ls.MoveFirst

For Kount = 0 To ls.RecordCount - 1
    List1.AddItem ls("guid")
    ls.MoveNext
Next

ls.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)

db.Close

End Sub
Private Sub List1_DblClick()
Text2.text = List1.List(List1.ListIndex)
End Sub

Private Sub Text2_Change()

'search the list box for what the user has typed

Dim Kount As Long
Dim Work As String
Dim TheOne As Long

TheOne = 0

For Kount = 0 To List1.ListCount - 1
    
    Work = UCase(List1.List(Kount))
    Work = Left(Work, Len(Text2.text))
    
    If Work = UCase(Text2.text) Then
        TheOne = Kount
    End If
    
Next

List1.ListIndex = TheOne
    
End Sub
