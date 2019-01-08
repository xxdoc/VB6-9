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
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtGrade 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblRemarks 
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblMarks 
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Marks"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Grade"
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
    
    Dim s As String
    
    s = txtGrade.Text
   
    Select Case s
        
        Case "E"
            lblMarks.Caption = "above 90%"
            lblRemarks.Caption = "Excellent"
        
        Case "A+"
            lblMarks.Caption = "above 80%"
            lblRemarks.Caption = "Very Good"
    
        Case "A"
            lblMarks.Caption = "above 70%"
            lblRemarks.Caption = "Good"
    
        Case "B"
            lblMarks.Caption = "above 60%"
            lblRemarks.Caption = "Average"
    
        Case "C"
            lblMarks.Caption = "above 50%"
            lblRemarks.Caption = "Satisfactory"
        
        Case "D"
            lblMarks.Caption = "above 40%"
            lblRemarks.Caption = "Poor"
        
        Case "F"
            lblMarks.Caption = "above 35%"
            lblRemarks.Caption = "Fail"
    
    End Select
    
End Sub
