VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()

    Dim word As String, n As Integer
    
    word = InputBox("Enter the word", "Input")
    n = Len(word)
    
    Select Case n
    
        Case 0
            Print "you have not entered any word"
        
        Case 1
            Print "This is a 1 letter word"
        
        Case 2
            Print "This is a 2 letter word"
        
        Case 3
            Print "This is a 3 letter word"
        
        Case 4
            Print "This is a 4 letter word"
        
        Case 5
            Print "This is a 5 letter word"
        
        Case Else
            Print "The word contains more than 5 letters"
    
    End Select
    
End Sub
