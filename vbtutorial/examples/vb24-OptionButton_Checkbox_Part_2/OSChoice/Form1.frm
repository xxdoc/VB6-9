VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox opt10 
      Caption         =   "Win 10"
      Height          =   1215
      Left            =   4080
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox opt8 
      Caption         =   "Win 8"
      Height          =   1215
      Left            =   2760
      Picture         =   "Form1.frx":3862
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox opt7 
      Caption         =   "Win 7"
      Height          =   1215
      Left            =   1440
      Picture         =   "Form1.frx":70C4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox optXp 
      Caption         =   "Win XP"
      Height          =   1215
      Left            =   120
      Picture         =   "Form1.frx":AB0E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optXp_Click()
    If optXp.Value = vbChecked Then
        MsgBox "You have chosen Windows XP", , "OS Selection"
    End If
End Sub

Private Sub opt7_Click()
    If opt7.Value = vbChecked Then
        MsgBox "You have chosen Windows 7", , "OS Selection"
    End If
End Sub

Private Sub opt8_Click()
    If opt8.Value = vbChecked Then
        MsgBox "You have chosen Windows 8", , "OS Selection"
    End If
End Sub

Private Sub opt10_Click()
    If opt10.Value = vbChecked Then
        MsgBox "You have chosen Windows 10", , "OS Selection"
    End If
End Sub
