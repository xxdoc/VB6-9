VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Program Maintenance"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Log Off"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reporting"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Inventory Maintenance"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Point Of Sale"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome!  Please choose an action from the list of items at the left."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2400
      TabIndex        =   5
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frmPOSMain.Visible = True
Unload Me

End Sub

Private Sub Command2_Click()

frmInventMaint.Visible = True
Unload Me

End Sub

Private Sub Command3_Click()

frmReports.Visible = True
Unload Me

End Sub

Private Sub Command5_Click()

Unload Me

End Sub

Private Sub Command6_Click()

frmSettings.Visible = True

End Sub

