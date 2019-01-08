VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMain"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtOutput 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtUpper 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "50"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtLower 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Upper Boundary"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Lower Boundary"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***
'*** Project:       Random Number Generator
'*** Subject:       VB6 Tutorial Lesson 47, example 1
'*** Date:          2018-06-06
'*** Author:        VB6boy
'***
'*** Purpose:       Demonstrates how to generate a random number, between a lower
'***                boundary and an upper boundary.
'***
Private Sub cmdGenerate_Click()

    Dim intLowerBoundary As Long            ' lower boundary
    Dim intUpperBoundary As Long            ' upper boundary
    Dim intOutput As Long                   ' output of random number generator
    
    Randomize Timer                         ' Seed the generator with the value from the system timer.
    
    intLowerBoundary = Int(txtLower.Text)   ' grab the lower limit
    intUpperBoundary = Int(txtUpper.Text)   ' grab the upper limit
    
    ' using the random number generator, get a random number with upper and
    ' lower limits, storing the result in intOutput
    
    intOutput = Int(Rnd(1) * intUpperBoundary) + intLowerBoundary
    
    txtOutput.Text = intOutput              ' send the output to the textbox
    
End Sub
