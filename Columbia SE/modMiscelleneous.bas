Attribute VB_Name = "modMiscelleneous"
Option Explicit

Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Function OpenLocation(Target As String, WindowState As Long) As Boolean

    Dim lHWnd As Long
    Dim lAns As Long
    
    lAns = ShellExecute(lHWnd, "open", Target, vbNullString, vbNullString, WindowState)
    OpenLocation = (lAns > 32)

End Function

Public Function LoadList(TargetListBox As ListBox, FileToLoad As String)

    '***
    '*** Load a listbox with the contents of the specified file
    '***
    
    Dim strWork As String
    
    Open FileToLoad For Input As #1
    While Not EOF(1)
        
        Line Input #1, strWork
        TargetListBox.AddItem strWork
        
    Wend
    Close #1
    
End Function

Public Function SaveList(SourceListBox As ListBox, FileToSave As String)

    '***
    '*** save the contents of a listbox to the specified file
    '***
    
    Dim lngKount As Long
    Dim strWork As String
    
    Open FileToSave For Output As #1
    For lngKount = 0 To SourceListBox.ListCount - 1
        strWork = Trim(SourceListBox.List(lngKount))
        Print #1, strWork
    Next
    Close #1
    
End Function
