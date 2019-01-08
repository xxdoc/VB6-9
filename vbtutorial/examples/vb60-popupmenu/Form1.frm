VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuColors 
      Caption         =   "Colors"
      Begin VB.Menu mnuBlue 
         Caption         =   "Blue"
      End
      Begin VB.Menu mnuGreen 
         Caption         =   "Green"
      End
      Begin VB.Menu mnuRed 
         Caption         =   "Red"
      End
      Begin VB.Menu mnuWhite 
         Caption         =   "White"
      End
      Begin VB.Menu mnuYellow 
         Caption         =   "Yellow"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
         PopupMenu mnuColors
    End If
End Sub

Private Sub mnuBlue_Click()
    Form1.BackColor = vbBlue
End Sub

Private Sub mnuGreen_Click()
    Form1.BackColor = vbGreen
End Sub

Private Sub mnuRed_Click()
    Form1.BackColor = vbRed
End Sub

Private Sub mnuWhite_Click()
    Form1.BackColor = vbWhite
End Sub

Private Sub mnuYellow_Click()
    Form1.BackColor = vbYellow
End Sub
