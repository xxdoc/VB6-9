VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Collumbia SE"
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   17670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddHost 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox txtHostToAdd 
      Height          =   285
      Left            =   5760
      TabIndex        =   10
      ToolTipText     =   "Enter a valid hostname in this box, and click the button to the right to add a host to the list"
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   6
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2520
      Width           =   7215
   End
   Begin VB.CommandButton cmdCheckSelected 
      Appearance      =   0  'Flat
      Caption         =   "Check &Selected"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton cmdPerformSweep 
      Appearance      =   0  'Flat
      Caption         =   "Check &All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   2175
   End
   Begin VB.ListBox lstHosts 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4320
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10920
      TabIndex        =   12
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblAddAHost 
      BackStyle       =   0  'Transparent
      Caption         =   "Add a Host:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Columbia SE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   11775
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Label lblHosts 
      BackStyle       =   0  'Transparent
      Caption         =   "Hosts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moveStartX As Integer
Dim moveStartY As Integer
Dim moveEndX As Integer
Dim moveEndY As Integer

Private Sub cmdAddHost_Click()

    Dim strHostToAdd As String
    
    strHostToAdd = Trim(txtHostToAdd.Text)
    If Len(strHostToAdd) < 1 Then
        Post txtOutput, vbCrLf & "!!! ERROR !!!" & vbCrLf & "!!! hostname to add cannot be blank !!!"
        Exit Sub
    Else
        lstHosts.AddItem strHostToAdd
        SaveList lstHosts, strHostFile
        Post txtOutput, strHostToAdd & " added, changes saved"
    End If
        
End Sub

Private Sub cmdCheckSelected_Click()

    Dim strHostToCheck As String
    
    strHostToCheck = lstHosts.List(lstHosts.ListIndex)
    ResolveAndPing strHostToCheck
     
End Sub

Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub cmdPerformSweep_Click()
    
    Dim Kount As Long
    Dim strHostName As String
    
    Post txtOutput, vbCrLf & "checking " & lstHosts.ListCount & " hosts ..." & vbCrLf
    
    For Kount = 0 To lstHosts.ListCount - 1
    
        strHostName = lstHosts.List(Kount)
        
        If Len(strHostName) < 1 Then
            
            Post txtOutput, "no hostname provided!"
            
        Else
        
            ResolveAndPing strHostName
                        
        End If
    
    Next
    
    Post txtOutput, vbCrLf & " cycle complete!" & vbCrLf
    
End Sub
Private Sub ResolveAndPing(strHostName As String)
            
    Dim Resolve As New clsResolver
    Dim strIPAddress As String
    Dim lngPingResult As Long
    Dim strPingResult As String
    
    Post txtOutput, "checking " & strHostName
    strIPAddress = Trim(Resolve.GetIPFromHostName(strHostName))
    
    If Len(strIPAddress) > 0 Then

        lngPingResult = DoPing(strIPAddress)
        
        Select Case lngPingResult
            
            Case 0
                strPingResult = "responding"
            Case 11001
                strPingResult = "buffer too small"
            Case 11002
                strPingResult = "destination network unreachable"
            Case 11003
                strPingResult = "destination host unreachable"
            Case 11004
                strPingResult = "destination protocol unreachable"
            Case 11005
                strPingResult = "destination port unreachable"
            Case 11006
                strPingResult = "no resources"
            Case 11007
                strPingResult = "bad option"
            Case 11008
                strPingResult = "hardware error"
            Case 11009
                strPingResult = "packet too big"
            Case 11010
                strPingResult = "request timed out"
            Case 11011
                strPingResult = "bad request"
            Case 11012
                strPingResult = "bad route"
            Case 11013
                strPingResult = "ttl expired in transit"
            Case 11014
                strPingResult = "ttl expired at reassembly"
            Case 11015
                strPingResult = "parameter problem"
            Case 11016
                strPingResult = "source quench"
            Case 11017
                strPingResult = "option too big"
            Case 11018
                strPingResult = "bad destination"
            Case 11032
                strPingResult = "negotiating ipsec"
            Case 11050
                strPingResult = "general failure"
            Case Else
                strPingResult = "unknowon return code"
        
        End Select
        
        Post txtOutput, "         " & strIPAddress & " " & strPingResult
        
    Else
        
        Post txtOutput, "         unable to resolve hostname"
    
    End If

End Sub

Private Sub Form_Load()

    Post txtOutput, "starting up ..."
    
    txtHostToAdd.Height = cmdAddHost.Height
    
    Post txtOutput, "default timeout set for " & pTimeOut
    Post txtOutput, "loading hosts ..."
    LoadList lstHosts, strHostFile
    Post txtOutput, lstHosts.ListCount & " hosts loaded"
    
    lstHosts.ToolTipText = "This is your host list. Single-click a host entry and click the 'Check Selected' button to check only the selected host. Double-click an entry to remove it from the list."
    
    lblClose.ToolTipText = "Click this to close the application"
    lblTitle.ToolTipText = "Thanks for trying Columbia SE. Why not try the full version?"
    lblHelp.ToolTipText = "Click here for assistance, or go to www.fortypoundhead.com for more information."
    
    txtHostToAdd.ToolTipText = "Enter a valid hostname in this box, and click the button to the right to add a host to the list"
    txtOutput.ToolTipText = "This scrolling text box will keep you informed as to what the current status is."
    
    cmdExit.ToolTipText = "Click this to close the application"
    cmdAddHost.ToolTipText = "After entering a hostname in the box to the left, click this to add it to the list. Your entry will be automatically saved."
    cmdPerformSweep.ToolTipText = "Check all the hosts in the list to see if they are responding to ICMP Echo Requests."
    cmdCheckSelected.ToolTipText = "Check the selected host to see if it is responding to ICMP Echo Requests."""
    
    Post txtOutput, "ready for action!" & vbCrLf
    
End Sub

Private Sub lblClose_Click()

    Unload Me
    
End Sub

Private Sub lblHelp_Click()

   Dim ret As Boolean
   ret = OpenLocation("https://www.fortypoundhead.com/newbrowseresults.asp?catid=102", SW_SHOWNORMAL)
    
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'You want to put these codes in your Forms_MouseDown event. This will keep track of
    'where the mouse is positioned when the button is clicked.
    '
    moveStartX = X
    moveStartY = Y
    
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Calculate the mouse position from its starting point to its current location.
    
    moveEndX = X - moveStartX
    moveEndY = Y - moveStartY
     
    If Button = 1 Then
     
        Me.Left = Me.Left + moveEndX
        Me.Top = Me.Top + moveEndY
     
    End If

End Sub

Private Sub lstHosts_Click()
    
    Dim strHostToCheck As String
    
    strHostToCheck = lstHosts.List(lstHosts.ListIndex)
    Post txtOutput, "selected " & strHostToCheck
    
    cmdCheckSelected.Enabled = True
    
End Sub

Private Sub lstHosts_DblClick()

    Dim strHostToRemove As String
    
    strHostToRemove = lstHosts.List(lstHosts.ListIndex)
    lstHosts.RemoveItem (lstHosts.ListIndex)
    SaveList lstHosts, strHostFile
    Post txtOutput, strHostToRemove & " removed, changes saved"
    
End Sub
