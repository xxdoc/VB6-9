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
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtNum3 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtNum2 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtNum1 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Number 3"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Number 2"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Number 1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim num1 As Integer, num2 As Integer, num3 As Integer
    num1 = Val(txtNum1.Text)
    num2 = Val(txtNum2.Text)
    num3 = Val(txtNum3.Text)
    
    If num1 > num2 And num1 > num3 Then
        lblResult.Caption = num1
    ElseIf num2 > num1 And num2 > num3 Then
        lblResult.Caption = num2
    Else
        lblResult.Caption = num3
    End If
End Sub
