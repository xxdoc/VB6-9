VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblMessage 
      Caption         =   "Write some text in the textbox:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPrint_Click()
    Dim strMyString As String
    strMyString = txtField.Text
    Print strMyString
End Sub

Private Sub cmdClear_Click()
    Cls       'Form1.Cls is also correct
End Sub

