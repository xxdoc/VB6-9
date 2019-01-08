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
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ASCII"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Character"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChange_Click()
    Text1.SetFocus

    If Label1.Caption = "Character" Then
        Label1.Caption = "ASCII"
        Label2.Caption = "Character"
    Else
        Label1.Caption = "Character"
        Label2.Caption = "ASCII"
    End If

    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Text1_Change()
    If Label1.Caption = "ASCII" Then
        Text2.Text = Chr(Val(Text1.Text))
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Label1.Caption = "Character" Then
        Text2.Text = KeyAscii
        Text1.Text = ""
    End If
End Sub
