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
   Begin VB.CommandButton cmdShow2 
      Caption         =   "Show2"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow1 
      Caption         =   "Show1"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox picOutput 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    picOutput.Cls       'Cls method clears the picture box
End Sub

Private Sub cmdShow1_Click()
    picOutput.Print "Hello World !"
End Sub

Private Sub cmdShow2_Click()
    picOutput.Print "Welcome to Visual Basic !"
End Sub
