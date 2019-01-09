VERSION 5.00
Begin VB.Form frmRptOutput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Output"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   11040
      TabIndex        =   1
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   12735
   End
End
Attribute VB_Name = "frmRptOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me

End Sub
