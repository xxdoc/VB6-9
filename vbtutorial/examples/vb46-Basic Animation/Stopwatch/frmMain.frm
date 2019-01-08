VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stopwatch"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHundredths 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Text            =   "00"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSeconds 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Text            =   "00"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtMinutes 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Text            =   "00"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtHours 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Text            =   "00"
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   2640
      Top             =   1440
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start / Stop"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***
'*** Project:       Stopwatch
'*** Subject:       VB6 Tutorial Lesson 46, example 2
'*** Date:          2018-05-08
'*** Author:        VB6boy
'***
'*** Purpose:       Demonstration of the timer control to update a digital
'***                clock. The clock is displayed in a textbox. The user is
'***                given the option of starting or pausing the clock update.
'***
'*** NOTE:          THIS IS ONLY A SIMULATION OF A STOPWATCH - ACCURACY IS NOT
'***                CODED INTO THIS EXAMPLE. IT IS ONLY MEANT FOR DEMONSTRATION
'***                OF UPDATING DISPLAY OBJECTS.
'***

'*** Define some variables

Dim intHours As Integer         ' hours
Dim intMinutes As Integer       ' minutes
Dim intSeconds As Integer       ' seconds
Dim intHundredths As Integer    ' hundredths of seconds
Dim bolIsRunning As Boolean     ' deterimines if the stopwatch is running or not

Private Sub cmdReset_Click()

'***
'*** Resets the values of the form-level variables, and calls for display update
'***

    intHours = 0
    intMinutes = 0
    intSeconds = 0
    intHundredths = 0
    
    UpdateDisplay
    
End Sub

Private Sub cmdStartStop_Click()

'***
'*** The start/stop method shown here rolls the functionality into one button
'*** Notice the enable/disable of the reset button here.  This disallows
'*** the user from clicking the reset button while the timer is running.
'***

    If bolIsRunning = False Then
        bolIsRunning = True
        tmrUpdate.Enabled = True
        cmdReset.Enabled = False
    Else
        bolIsRunning = False
        tmrUpdate.Enabled = False
        cmdReset.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()

'*** This is the initial form load, where some starting points will be set.

    bolIsRunning = False        ' turn off the running switch
    tmrUpdate.Enabled = False   ' ensure the timer is turned off
    tmrUpdate.Interval = 10     ' Only tracking 100ths of a second, so update every 5 ms

    UpdateDisplay
    
End Sub

Private Sub tmrUpdate_Timer()

    intHundredths = intHundredths + 2
    
    If intHundredths > 99 Then
        intHundredths = 0
        intSeconds = intSeconds + 1
    End If
    
    If intSeconds > 59 Then
        intSeconds = 0
        intMinutes = intMinutes + 1
    End If
    
    If intMinutes > 59 Then
        intMinutes = 0
        intHours = intHours + 1
    End If
    
    UpdateDisplay
        
End Sub

Private Sub UpdateDisplay()

    txtHours.Text = Format(intHours, "00")
    txtMinutes.Text = Format(intMinutes, "00")
    txtSeconds.Text = Format(intSeconds, "00")
    txtHundredths.Text = Format(intHundredths, "00")
    
End Sub
