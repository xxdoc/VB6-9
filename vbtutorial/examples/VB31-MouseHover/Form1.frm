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
   Begin VB.CommandButton cmdClickMe 
      Caption         =   "Click Me"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtMessage 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape shpClickMe 
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   960
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClickMe_Click()
    
    txtMessage.Text = "Welcome to the site!"
    
End Sub

Private Sub cmdClickMe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    shpClickMe.Visible = True
    cmdClickMe.FontSize = 17
    cmdClickMe.BackColor = vbGreen

End Sub

Private Sub Form_Load()

    txtMessage.Text = ""
    shpClickMe.Visible = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    shpClickMe.Visible = False
    cmdClickMe.FontSize = 10
    cmdClickMe.BackColor = &H8000000F
    
End Sub
