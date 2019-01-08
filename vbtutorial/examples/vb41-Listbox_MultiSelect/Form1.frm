VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload Listbox"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   5655
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Selected"
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeleteAll 
      Caption         =   "Delete All"
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClearAll_Click()
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = False
    Next
End Sub

Private Sub cmdDelete_Click()
    For i = 0 To List1.SelCount - 1
        List1.RemoveItem List1.ListIndex
    Next
End Sub

Private Sub cmdDeleteAll_Click()
    List1.Clear
End Sub

Private Sub cmdReload_Click()
    LoadListbox List1
End Sub

Private Sub cmdSelectAll_Click()
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = True
    Next
End Sub

Private Sub Form_Load()
    LoadListbox List1
End Sub

Private Sub LoadListbox(MyList As ListBox)
    MyList.Clear
    For i = 0 To 10
        MyList.AddItem "Item" & i
    Next
End Sub
