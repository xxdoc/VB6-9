VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox chkCoffee 
      Caption         =   "Coffee $9"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox chkGreenTea 
      Caption         =   "Green Tea $7"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkColdDrinks 
      Caption         =   "Cold Drinks $6"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CheckBox chkPizza 
      Caption         =   "Pizza $50"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox chkSandwich 
      Caption         =   "Sandwich $50"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkTea 
      Caption         =   "Tea $5"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdShow_Click()
    
    Dim total As Integer
    
    total = 0
    tea = ""
    sandwitch = ""
    pizza = ""
    colddrinks = ""
    greentea = ""
    coffee = ""


    If chkTea.Value = 1 Then
        tea = " Tea,"
        total = total + 5
    End If

    If chkSandwich.Value = 1 Then
        sandwich = " Sandwich,"
        total = total + 50
    End If

    If chkPizza.Value = 1 Then
        pizza = " Pizza,"
        total = total + 50
    End If

    If chkColdDrinks.Value = 1 Then
        colddrinks = " Cold Drinks,"
        total = total + 6
    End If

    If chkGreenTea.Value = 1 Then
        greentea = " Green Tea,"
        total = total + 7
    End If

    If chkCoffee.Value = 1 Then
        coffee = " Coffee,"
        total = total + 9
    End If


    MsgBox "You have ordered " & tea & _
    sandwich & pizza & colddrinks & greentea & coffee & _
    " and the total price is $" & total, vbInformation, "Thanks"


End Sub

