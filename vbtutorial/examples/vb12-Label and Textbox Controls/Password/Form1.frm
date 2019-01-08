VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Note: Password is 12345 but try an incorrect password to see what happens"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Password"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirm_Click()
    If txtField.Text = "12345" Then
        MsgBox "Successful !!!", vbInformation, ""
    Else
        MsgBox "Incorrect Password !", vbCritical, ""
    End If
End Sub

Private Sub cmdClear_Click()
    txtField.Text = ""
End Sub
