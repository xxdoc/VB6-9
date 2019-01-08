Attribute VB_Name = "modStartup"
Option Explicit
Global pTimeOut As Long
Global strHostFile As String

Sub main()

    strHostFile = App.Path & "\hosts.txt"
    pTimeOut = 200
    
    frmMain.Visible = True
    
End Sub
