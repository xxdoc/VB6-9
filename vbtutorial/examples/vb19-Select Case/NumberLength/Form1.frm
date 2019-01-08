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
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtNum 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblResult 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter A Number"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()

    Dim n As Long
    
    n = Val(txtNum.Text)
   
    Select Case n
    
        Case 0 To 9
            lblResult.Caption = "Single digit number"
           
        Case 10 To 99
            lblResult.Caption = "two digit number"
    
        Case 100 To 999
            lblResult.Caption = "Three digit number"
    
        Case 1000 To 9999
             lblResult.Caption = "Four digit number"
    
        Case 10000 To 99999
            lblResult.Caption = "Five digit number"
    
        Case Else
            lblResult.Caption = "More than Five digit number"
        
    End Select
    
End Sub
