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
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtResult 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtValue2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtValue1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Sum"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Value 2"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Value 1"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim a As Integer, b As Integer
    a = Val(txtValue1.Text)  'val function converts string
                             'into value
    b = Val(txtValue2.Text)
    txtResult.Text = a + b
End Sub

Private Sub cmdClear_Click()
    txtValue1.Text = ""
    txtValue2.Text = ""
    txtResult.Text = ""
End Sub

