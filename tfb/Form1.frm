VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim strRandomString As String
    Dim strWork As String
    Dim intRandomASCII As Integer
    Dim lngKount As Long
    Dim BigLoop As Currency
    
    Open App.Path & "\output.txt" For Output As #1
    
    For BigLoop = 1 To 100000
        
        strRandomString = ""
        
        For lngKount = 1 To 1024
            intRandomASCII = Int(Rnd(1) * 50) + 65
            strRandomString = strRandomString & Chr(intRandomASCII)
        Next
    
        Print #1, strRandomString
        Debug.Print BigLoop
        DoEvents
        
    Next
    
    Close #1
    
    
    
End Sub

Private Sub Command2_Click()

    Open App.Path & "\o3.txt" For Input As #1
    Open App.Path & "\o4.txt" For Append As #2
    
    While Not EOF(1)
        Line Input #1, strWork
        Print #2, strWork
    Wend
    
    Close #2
    Close #1
    
    Debug.Print "done"
    
End Sub
