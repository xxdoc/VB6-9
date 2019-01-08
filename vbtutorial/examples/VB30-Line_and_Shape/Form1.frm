VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line8 
      X1              =   120
      X2              =   5160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line7 
      X1              =   5160
      X2              =   5160
      Y1              =   720
      Y2              =   120
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   120
      Y1              =   720
      Y2              =   120
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4815
   End
   Begin VB.Line Line5 
      X1              =   1800
      X2              =   1800
      Y1              =   720
      Y2              =   1920
   End
   Begin VB.Line Line4 
      X1              =   5160
      X2              =   5160
      Y1              =   1920
      Y2              =   720
   End
   Begin VB.Line Line3 
      X1              =   5160
      X2              =   120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   1920
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblHeader 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblData 
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    lblCompany.Caption = "Initrode"
    lblHeader.Caption = "Name:" & vbCrLf & "Company:" & vbCrLf & "Department:" & vbCrLf & "Team:" & vbCrLf & "ID:"
    lblData.Caption = "John Doe" & vbCrLf & "Initrode" & vbCrLf & "Information Technology" & vbCrLf & "Automation" & vbCrLf & "1234567890-432"

End Sub
