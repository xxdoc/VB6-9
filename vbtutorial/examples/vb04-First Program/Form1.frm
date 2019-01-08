VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "My First VB Program"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()                    ' this is a sample program
    Print "Welcome to www.fortypoundhead.com"   ' this is a demo program for beginners
    ' you can also put comments on a line by themselves
End Sub
