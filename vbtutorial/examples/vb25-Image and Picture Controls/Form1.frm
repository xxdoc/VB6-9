VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShow_Click()
    Picture1.Picture = LoadPicture(txtPath.Text)
End Sub
