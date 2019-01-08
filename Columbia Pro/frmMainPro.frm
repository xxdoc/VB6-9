VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMainPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMainPro"
   ClientHeight    =   5010
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPingSelected 
      Caption         =   "Ping Selected"
      Height          =   495
      Left            =   10200
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdPingAll 
      Caption         =   "Ping All"
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      Height          =   3975
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   360
      Width           =   5655
   End
   Begin VB.FileListBox filGroups 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.ListBox lstGroupMembers 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog cdlgOpenSave 
      Left            =   120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblMembers 
      Caption         =   "Group Members:"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblGroups 
      Caption         =   "Groups:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGroup 
         Caption         =   "&New Group"
      End
      Begin VB.Menu mnuSaveGroup 
         Caption         =   "&Save Group"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuStartPingAll 
         Caption         =   "&Ping All"
      End
      Begin VB.Menu mnuPingSelected 
         Caption         =   "Ping &Selected"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuVisitHelpPages 
         Caption         =   "&Visit Help Pages"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMainPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************
'***
'*** MODULE: frmMainPro
'*** AUTHOR: Derek A. Wirch
'*** DATE: 07/15/2018
'***
'*** PURPOSE: Main interface
'***
'*** SIDE EFFECTS: None
'***
'***************************************************************************************
'***
'*** HISTORY:
'*** Date Description
'*** ---------- --------------------------------------------------------------------
'*** 07/15/2018 Initial Creation
'*** 07/18/2018 added disable/enable controls for ping checks
'***            added new directory (source) for html source code storage
'***
'***************************************************************************************

Dim strSelectedHost As String       ' the currently selected host
Dim strSelectedGroup As String      ' the currently selected group

Private Sub cmdPingAll_Click()

    '***
    '*** Redirector to menu entry
    '***
    
    mnuStartPingAll_Click
    
End Sub

Private Sub cmdPingSelected_Click()

    '***
    '*** Redirector to menu entry
    '***

    mnuPingSelected_Click

End Sub

Private Sub filGroups_Click()
'***************************************************************************************
'***
'*** FUNCTION: filGroups_Click
'*** AUTHOR: Derek A. Wirch
'*** DATE: 07/15/2018
'***
'*** PURPOSE: Loads the contents of the selected group
'*** RETURNS: None
'*** USAGE: Initiated by clicking group name in the listbox
'***
'*** SIDE EFFECTS: None
'***
'***************************************************************************************
'***
'*** HISTORY:
'*** Date Description
'*** ---------- --------------------------------------------------------------------
'*** 07/15/2018 Initial Creation
'***
'***************************************************************************************
    
    Dim strWork As String
    
    strSelectedGroup = Trim(filGroups.List(filGroups.ListIndex))
    Post txtStatus, "loading " & strSelectedGroup
    lstGroupMembers.Clear
    strSelectedHost = ""
    
    Open GroupsDir & strSelectedGroup For Input As #1
    While Not EOF(1)
        Line Input #1, strWork
        lstGroupMembers.AddItem strWork
    Wend
    Close #1
    
    Post txtStatus, "loaded " & strSelectedGroup, True
        
End Sub

Private Sub Form_Load()

'***************************************************************************************
'***
'*** FUNCTION: Form_Load
'*** AUTHOR: Derek A. Wirch
'*** DATE: 07/15/2018
'***
'*** PURPOSE: form initialization event
'*** RETURNS: None
'*** USAGE: fired from sub_main
'***
'*** SIDE EFFECTS: None
'***
'***************************************************************************************
'***
'*** HISTORY:
'*** Date Description
'*** ---------- --------------------------------------------------------------------
'*** 07/15/2018 Initial Creation
'***
'***************************************************************************************

    '***
    '*** set up file list box
    '***
    
    filGroups.Path = GroupsDir
    filGroups.Pattern = "*."
    filGroups.Refresh
    
End Sub

Private Sub lstGroupMembers_Click()

'***************************************************************************************
'***
'*** FUNCTION: lstGroupMembers_Click
'*** AUTHOR: Derek A. Wirch
'*** DATE: 07/15/2018
'***
'*** PURPOSE: populate the selected host variable with date user clicked on.
'*** RETURNS: None
'*** USAGE: fired on click of host list
'***
'*** SIDE EFFECTS: None
'***
'***************************************************************************************
'***
'*** HISTORY:
'*** Date Description
'*** ---------- --------------------------------------------------------------------
'*** 07/15/2018 Initial Creation
'***
'***************************************************************************************
    
    strSelectedHost = lstGroupMembers.List(lstGroupMembers.ListIndex)
    strSelectedHost = Trim(strSelectedHost)
    Post txtStatus, "selected " & strSelectedHost, True

End Sub

Private Sub mnuPingSelected_Click()

'***************************************************************************************
'***
'*** FUNCTION: mnuPingSelected_Click
'*** AUTHOR: Derek A. Wirch
'*** DATE: 07/15/2018
'***
'*** PURPOSE: Pings the selected host
'*** RETURNS: None
'*** USAGE: fired on button click/menu click
'***
'*** SIDE EFFECTS: None
'***
'***************************************************************************************
'***
'*** HISTORY:
'*** Date Description
'*** ---------- --------------------------------------------------------------------
'*** 07/15/2018 Initial Creation
'***
'***************************************************************************************
    
    Dim strPingResult As String
    Dim strHistoryFile As String
    
    If strSelectedHost = "" Then
        
        MsgBox "You must select a host before attempting a ping check", vbExclamation, "No host selected!"
        Post txtStatus, "no host selected", True
    
    Else
    
        Post txtStatus, "attempting to resolve and ping " & strSelectedHost, True
        DisablePingButtons
        
        strPingResult = ResolveAndPing(strSelectedHost)
        Post txtStatus, strPingResult, True
        
        '***
        '*** note results in history
        '***
        
        strHistoryFile = HistDir & uDate & "_" & strSelectedHost & ".log"
        Open strHistoryFile For Append As #1
        Print #1, uDateTime & " " & strPingResult
        Close #1
        EnablePingButtons
        
    End If
    
End Sub

Private Sub mnuStartPingAll_Click()

'***************************************************************************************
'***
'*** FUNCTION: mnuStartPingAll_Click
'*** AUTHOR: Derek A. Wirch
'*** DATE: 07/15/2018
'***
'*** PURPOSE: Pings all hosts in the selected group
'*** RETURNS: None
'*** USAGE: fired on button click/menu click
'***
'*** SIDE EFFECTS: None
'***
'***************************************************************************************
'***
'*** HISTORY:
'*** Date Description
'*** ---------- --------------------------------------------------------------------
'*** 07/15/2018 Initial Creation
'***
'***************************************************************************************
    
    Dim strPingResult As String
    Dim strHistoryFile As String
    Dim lngKount As Long
    Dim strCurrentHost As String
    Dim intMyFile As Integer
    
    If strSelectedGroup = "" Then
        
        MsgBox "You must select a group to ping before attempting this", vbExclamation, "No group selected!"
        Post txtStatus, "no group selected", True
        
    Else
        
        Post txtStatus, "beginning check of " & strSelectedGroup, True
        DisablePingButtons
        
        For lngKount = 0 To lstGroupMembers.ListCount - 1
            
            strCurrentHost = Trim(lstGroupMembers.List(lngKount))
            strHistoryFile = HistDir & uDate & "_" & strCurrentHost & ".log"

            Post txtStatus, "attempting to resolve and ping " & strCurrentHost
            strPingResult = ResolveAndPing(strCurrentHost)
            Post txtStatus, strPingResult
            
            intMyFile = FreeFile()
            
            Open strHistoryFile For Append As #intMyFile
            Print #intMyFile, uDateTime & " " & strPingResult
            Close #intMyFile

        Next
        
        EnablePingButtons
        Post txtStatus, "group check complete for " & strSelectedGroup, True
        
    End If
    
End Sub

Private Sub DisablePingButtons()

'***************************************************************************************
'***
'*** FUNCTION: DisablePingButtons
'*** AUTHOR: Derek A. Wirch
'*** DATE: 07/15/2018
'***
'*** PURPOSE: Turns off buttons and menu entries (while ping check in progress)
'*** RETURNS: None
'*** USAGE: fired when ping cycle started
'***
'*** SIDE EFFECTS: None
'***
'***************************************************************************************
'***
'*** HISTORY:
'*** Date Description
'*** ---------- --------------------------------------------------------------------
'*** 07/15/2018 Initial Creation
'***
'***************************************************************************************

    cmdPingAll.Enabled = False
    cmdPingSelected.Enabled = False
    mnuStartPingAll.Enabled = False
    mnuPingSelected.Enabled = False

End Sub

Private Sub EnablePingButtons()

'***************************************************************************************
'***
'*** FUNCTION: EnablePingButtons
'*** AUTHOR: Derek A. Wirch
'*** DATE: 07/15/2018
'***
'*** PURPOSE: Turns on the ping buttons and menu entries
'*** RETURNS: None
'*** USAGE: fired when ping cycle complete
'***
'*** SIDE EFFECTS: None
'***
'***************************************************************************************
'***
'*** HISTORY:
'*** Date Description
'*** ---------- --------------------------------------------------------------------
'*** 07/15/2018 Initial Creation
'***
'***************************************************************************************
    
    cmdPingAll.Enabled = True
    cmdPingSelected.Enabled = True
    mnuStartPingAll.Enabled = True
    mnuPingSelected.Enabled = True

End Sub
