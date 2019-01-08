VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFocusDemo02 
      Caption         =   "Command2"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdFocusDemo01 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFocusDemo01_GotFocus()

    Debug.Print "Command 1 has got the focus"
    
End Sub

Private Sub cmdFocusDemo01_LostFocus()

    Debug.Print "Command 1 has lost the focus"
    
End Sub

Private Sub cmdFocusDemo02_GotFocus()

    Debug.Print "Command 2 has got the focus"
    
End Sub

Private Sub cmdFocusDemo02_LostFocus()

    Debug.Print "Command 2 has lost the focus"
    
End Sub
