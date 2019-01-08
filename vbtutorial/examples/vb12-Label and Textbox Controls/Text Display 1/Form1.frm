VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtField 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label lblMessage 
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShow_Click()
    txtField.Text = "Hello World"
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdClear_Click()
    txtField.Text = ""
End Sub
