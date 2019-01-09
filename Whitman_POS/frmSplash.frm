VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "frmSplash"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6240
      Top             =   1440
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10920
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   8640
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   10695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   2490
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Label1.Caption = App.Title
    Label2(0).Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    Label2(1).Caption = App.CompanyName
    Label2(2).Caption = App.LegalCopyright
    
    Line1.Y1 = Image1.Top + Image1.Height
    Line1.Y2 = Image1.Top + Image1.Height


End Sub

Private Sub Image1_Click()

    Unload Me

End Sub

Private Sub Label1_Click()

    Unload Me

End Sub

Private Sub Label2_Click(Index As Integer)

    Unload Me

End Sub

Private Sub Timer1_Timer()

    Unload Me

End Sub
