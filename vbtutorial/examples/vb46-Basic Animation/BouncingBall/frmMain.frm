VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bouncing Ball"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8340
   FillColor       =   &H0000FF00&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start/Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin VB.Shape shpBall 
      BackStyle       =   1  'Opaque
      BorderStyle     =   2  'Dash
      FillStyle       =   7  'Diagonal Cross
      Height          =   495
      Left            =   600
      Shape           =   1  'Square
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngX As Long                ' current horizontal postion of the top/left corner of the ball
Dim lngY As Long                ' current vertical position of the top/left corner of the ball
Dim intBallWidth As Integer     ' width of the ball (width and height should be the same)
Dim intBallHeight As Integer    ' height of the ball (width and height should be the same)
Dim intXStep As Integer         ' step size for horizontal
Dim intYstep As Integer         ' step size for vertical
Dim bolMovingDown As Boolean    ' moving up or down
Dim bolMovingRight As Boolean   ' moving left or right

Private Sub cmdStartStop_Click()

    If tmrUpdate.Enabled = True Then
        tmrUpdate.Enabled = False
    Else
        tmrUpdate.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()

'*** set some initial values

    lngX = 100
    lngY = 100
    intBallWidth = 500
    intBallHeight = 500
    
    tmrUpdate.Enabled = False
    tmrUpdate.Interval = 50
        
End Sub

Private Sub tmrUpdate_Timer()

    lngX = lngX + intXStep
    lngY = lngY + intYstep
    
    If lngX + intBallWidth > Me.Width Then
        
        intXStep = -intXStep
        lngX = lngX + intXStep
        
    End If
    
    If lngY + intBallHeight > Me.Height Then
        
        intYstep = -intYstep
        lngY = lngY + intYstep
        
    End If
    
    '*** update position
    
    shpBall.Top = lngY
    shpBall.Left = lngX
    
End Sub
