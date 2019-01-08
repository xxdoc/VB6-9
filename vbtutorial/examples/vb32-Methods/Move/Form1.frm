VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdTopRightCorner 
      Caption         =   "Move up and Right"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetBigger 
      Caption         =   "Get Bigger"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdMoveToRight 
      Caption         =   "Move Right"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' these are for the default values of the textbox controls

Dim lngText1Top As Long
Dim lngText1Left As Long

Dim lngText2Width As Long
Dim lngText2Height As Long
Dim lngText2Top As Long
Dim lngText2Left As Long

Dim lngText3Top As Long
Dim lngText3Left As Long

' and this one is the default increment that we will move the controls

Dim lngIncrement As Long

Private Sub ResetControls()
    
    ' reset the controls to their original values,
    ' using the values we captured in the form_load event.
    
    Text1.Left = lngText1Left
    Text1.Top = lngText1Top
    
    Text2.Top = lngText2Top
    Text2.Left = lngText2Left
    Text2.Width = lngText2Width
    Text2.Height = lngText2Height
    
    Text3.Top = lngText3Top
    Text3.Left = lngText3Left
    
End Sub

Private Sub cmdMoveToRight_Click()

    Text1.Move Text1.Left + lngIncrement

End Sub

Private Sub cmdGetBigger_Click()

    Text2.Move lngText2Left, lngText2Top, Text2.Width + lngIncrement, Text2.Height + lngIncrement
    
End Sub

Private Sub cmdReset_Click()

    ResetControls
    
End Sub

Private Sub cmdTopRightCorner_Click()

    If (Text3.Top - lngIncrement) > 0 Then
        
        ' We only want to move the control up to 100 or so. if it
        ' gets below zero, errors will ensue, crashing our program.
        
        Text3.Move Text3.Left + lngIncrement, Text3.Top - lngIncrement
    
    End If
    
    
End Sub

Private Sub Form_Load()

    ' When the form loads, we are using the module level
    ' long integer variables, defined at the top of this
    ' module. This allows us to restore the original
    ' positions in the ResetControls subroutine.
    
    lngText1Top = Text1.Top
    lngText1Left = Text1.Left
    
    lngText2Height = Text2.Height
    lngText2Width = Text2.Width
    lngText2Top = Text2.Top
    lngText2Left = Text2.Left
    
    lngText3Top = Text3.Top
    lngText3Left = Text3.Left
    
    lngIncrement = 100
        
End Sub
