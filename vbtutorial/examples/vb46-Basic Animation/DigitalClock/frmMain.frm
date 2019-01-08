VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Clock"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdateClock 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   1560
   End
   Begin VB.TextBox txtClock 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Text            =   "Click Start"
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***
'*** Project:       Digital Clock
'*** Subject:       VB6 Tutorial Lesson 46, example 1
'*** Date:          2018-05-08
'*** Author:        VB6boy
'***
'*** Purpose:       Demonstration of the timer control to update a digital
'***                clock. The clock is displayed in a textbox. The user is
'***                given the option of starting or pausing the clock update.
'***
'***
'***

Private Sub cmdPause_Click()

'***
'*** The user clicked the Pause button, so disable the timer, stopping the
'*** update of the display when the timer fires.
'***

    tmrUpdateClock.Enabled = False
    
End Sub
Private Sub cmdStart_Click()

'***
'*** The user has clicked the start button, so enable the timer, restarting the
'*** update of the display when the timer fires.
'***

    tmrUpdateClock.Enabled = True
    
End Sub

Private Sub tmrUpdateClock_Timer()

'***
'*** Timer event fires, updates the textbox with the current time.
'***

    txtClock.Text = Time
    
End Sub
