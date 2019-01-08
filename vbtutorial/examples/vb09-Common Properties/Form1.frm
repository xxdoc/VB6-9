VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdChangePosition 
      Caption         =   "Change Position"
      Height          =   735
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdMakeSmaller 
      Caption         =   "Shrink"
      Height          =   735
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnlarge 
      Caption         =   "Enlarge"
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdWhite 
      Caption         =   "White"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdBlack 
      Caption         =   "Black"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdGreen 
      Caption         =   "Green"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "Red"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBlack_Click()
    Form1.BackColor = vbBlack
End Sub

Private Sub cmdChangePosition_Click()
    Form1.Top = 5000
    Form1.Left = 5000
End Sub

Private Sub cmdDefault_Click()
    Form1.BackColor = &H8000000F
    Form1.Width = 4500
    Form1.Height = 4125
    Form1.Top = 900
    Form1.Left = 900
End Sub

Private Sub cmdEnlarge_Click()
    Form1.Width = 5160
    Form1.Height = 4770
End Sub

Private Sub cmdGreen_Click()
    Form1.BackColor = vbGreen
End Sub

Private Sub cmdMakeSmaller_Click()
    Form1.Width = 4500
    Form1.Height = 4125
End Sub

Private Sub cmdRed_Click()
    Form1.BackColor = vbRed
End Sub

Private Sub cmdWhite_Click()
    Form1.BackColor = vbWhite
End Sub

Private Sub Form_Load()
    Form1.Top = 900
    Form1.Left = 900
End Sub
