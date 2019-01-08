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
   Begin VB.Frame fraTextColor 
      Caption         =   "Text Color"
      Height          =   1935
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
      Begin VB.OptionButton optGreen 
         Caption         =   "Green"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "Blue"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optRed 
         Caption         =   "Red"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraFontSize 
      Caption         =   "Font Size"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
      Begin VB.OptionButton optLarge 
         Caption         =   "Large"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optMedium 
         Caption         =   "Medium"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optSmall 
         Caption         =   "Small"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label lblMessage 
      Caption         =   "Welcome to Visual Basic 6"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optBlue_Click()
    lblMessage.ForeColor = vbBlue
End Sub

Private Sub optGreen_Click()
    lblMessage.ForeColor = vbGreen
End Sub

Private Sub optLarge_Click()
    lblMessage.FontSize = 22
End Sub

Private Sub optMedium_Click()
    lblMessage.FontSize = 16
End Sub

Private Sub optRed_Click()
    lblMessage.ForeColor = vbRed
End Sub

Private Sub optSmall_Click()
    lblMessage.FontSize = 12
End Sub
