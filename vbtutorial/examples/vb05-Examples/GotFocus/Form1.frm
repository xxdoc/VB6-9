VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Use the tab key to move the focus between the buttons"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_GotFocus()

    ' update label2
    
    Label2.Caption = "Command1 has the focus"
    
End Sub
Private Sub Command2_GotFocus()

    ' update label2
    
    Label2.Caption = "Command2 has the focus"
    
End Sub
Private Sub Command3_GotFocus()

    ' update label2
    
    Label2.Caption = "Command3 has the focus"
    
End Sub
Private Sub Command4_GotFocus()

    ' update label2
    
    Label2.Caption = "Command4 has the focus"
    
End Sub
Private Sub Command5_GotFocus()

    ' update label2
    
    Label2.Caption = "Command5 has the focus"
    
End Sub
Private Sub Command6_GotFocus()

    ' update label2
    
    Label2.Caption = "Command6 has the focus"
    
End Sub
