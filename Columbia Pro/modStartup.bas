Attribute VB_Name = "modStartup"
Option Explicit

Global GroupsDir As String
Global LogDir As String
Global HistDir As String
Global SourceDir As String

Global pTimeOut As Long


Sub main()

    pTimeOut = 200
    
    GroupsDir = App.Path & "\groups\"
    LogDir = App.Path & "\logs\"
    HistDir = App.Path & "\history\"
    SourceDir = App.Path & "\source\"
    
    SendToLogFile "Startup complete:" & vbCrLf & _
                  Space(20) & "GroupsDir   " & GroupsDir & vbCrLf & _
                  Space(20) & "LogDir      " & LogDir & vbCrLf & _
                  Space(20) & "HistDir     " & HistDir & vbCrLf & _
                  Space(20) & "SourceDir   " & SourceDir & vbCrLf & _
                  Space(20) & "Timeout     " & pTimeOut & vbCrLf & _
                  Space(20) & "Firing frmMain"
                  
    frmMainPro.Visible = True
    
End Sub
