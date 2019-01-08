VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Length Converter"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCent 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtInch 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3015
      LargeChange     =   10
      Left            =   120
      Max             =   1000
      Min             =   1
      TabIndex        =   0
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Inches"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Centimeters"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub VScroll1_Change()

    txtInch.Text = VScroll1.Value
    txtCent.Text = VScroll1.Value * 2.54
    
End Sub

Private Sub VScroll1_Scroll()

    txtInch.Text = VScroll1.Value
    txtCent.Text = VScroll1.Value * 2.54
    
End Sub
