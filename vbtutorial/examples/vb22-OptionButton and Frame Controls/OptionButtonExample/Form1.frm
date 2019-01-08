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
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.OptionButton optCricket 
      Caption         =   "Cricket"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.OptionButton optFootball 
      Caption         =   "Football"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtOutput 
      Height          =   2895
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirm_Click()
    
    If optFootball.Value = True Then
        txtOutput.Text = "You have selected football."
    ElseIf optCricket.Value = True Then
        txtOutput.Text = "You have selected cricket."
    End If
    
End Sub
