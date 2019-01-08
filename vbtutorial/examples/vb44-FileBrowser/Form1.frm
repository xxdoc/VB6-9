VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCurrentPath 
      Height          =   285
      Left            =   3225
      TabIndex        =   4
      Top             =   5520
      Width           =   9270
   End
   Begin VB.PictureBox picDisplay 
      Height          =   4935
      Left            =   6240
      ScaleHeight     =   4875
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   480
      Width           =   6255
   End
   Begin VB.FileListBox filSelect 
      Height          =   4965
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.DirListBox dirSelect 
      Height          =   5040
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.DriveListBox drvSelect 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Currently Selected Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   5550
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFullPath As String           ' module-wide string to hold the currently selected path

Private Sub dirSelect_Change()

    ' When the directory selection changes, we must also update the file selection path
    
    filSelect.Path = dirSelect.Path
    UpdatePath
    
End Sub

Private Sub drvSelect_Change()

    ' when the drive selection changes, we must also update the directory selection path
    
    dirSelect.Path = drvSelect.Drive
    UpdatePath
    
End Sub

Private Sub filSelect_Click()
    
    ' When the user clicks on a file in the file list, this will load the selected picture to the picturebox.
    ' with a little math, and the use of a couple of properties stolen from the properties of both the picturebox
    ' and the picture itself, the picture is scaled down or up to fit the picture box. However, there is still
    ' distortion, due to not taking into account the aspect ration of the picture vs. the dimensions of the
    ' picturebox. For example, 16:9 widescreen pictures will appear horizontally "squashed", while 4:3 pictures
    ' will appear near normal.
    
    ' to do: fix the aspect ration problem
    
    picDisplay.Picture = LoadPicture(strFullPath & "\" & filSelect.FileName)
    picDisplay.ScaleMode = 3
    picDisplay.AutoRedraw = True
    picDisplay.PaintPicture picDisplay.Picture, 0, 0, picDisplay.ScaleWidth, picDisplay.ScaleHeight, 0, 0, picDisplay.Picture.Width / 26.46, picDisplay.Picture.Height / 26.46
    picDisplay.Picture = picDisplay.Image
   
End Sub

Private Sub Form_Load()
    
    ' set the file display pattern to the types of files that we want to see in the file selector box
    
    filSelect.Pattern = "*.jpg;*.jpeg;*.bmp"
    
End Sub

Private Sub UpdatePath()

    ' this does nothing but update the path variable, as well as display the currently selected path in the
    ' text field at the bottom of the window.
    
    strFullPath = dirSelect.Path
    txtCurrentPath.Text = strFullPath
    
End Sub
