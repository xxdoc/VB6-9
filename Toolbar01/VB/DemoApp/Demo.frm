VERSION 5.00
Begin VB.Form DemoBar 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "AppBar Demo"
   ClientHeight    =   390
   ClientLeft      =   3000
   ClientTop       =   1980
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnProperties 
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FF00FF&
      Picture         =   "Demo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Properties"
      Top             =   15
      UseMaskColor    =   -1  'True
      Width           =   375
   End
End
Attribute VB_Name = "DemoBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AppBar As New TAppBar

Private Sub Form_Load()
  
  AppBar.Extends Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  AppBar.Detach

End Sub

Private Sub btnProperties_Click()
  
  PropDlg.Show vbModal
  AppBar.UpdateBar

End Sub

