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
   Begin VB.CommandButton cmdEnterValues 
      Caption         =   "Enter Values"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnterValues_Click()
    a = Val(InputBox("Enter the value of a", "Input value"))
    b = Val(InputBox("Enter the value of b", "Input value"))
    MsgBox "The sum is " & a + b, vbInformation, "Output"
End Sub
