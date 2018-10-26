'In a Class Module Named UndoElement

Public SelStart As Long
Public TextLen As Long
Public Text As String


'In a Module
Public trapUndo As Boolean
Public UndoStack As New Collection
Public RedoStack As New Collection
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Public Sub Redo()
Dim chg$
Dim DeleteFlag As Boolean
Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(RTFForm.RichTextBox.Text)
        If DeleteFlag Then
            Set objElement = RedoStack(RedoStack.Count)
            RTFForm.RichTextBox.SelStart = objElement.SelStart
            RTFForm.RichTextBox.SelLength = Len(RTFForm.RichTextBox.Text) - objElement.TextLen
            RTFForm.RichTextBox.SelText = ""
        Else
            Set objElement = RedoStack(RedoStack.Count)
            chg$ = Change(RTFForm.RichTextBox.Text, objElement.Text, objElement.SelStart + 1)
            RTFForm.RichTextBox.SelStart = objElement.SelStart - Len(chg$)
            RTFForm.RichTextBox.SelLength = 0
            RTFForm.RichTextBox.SelText = chg$
            RTFForm.RichTextBox.SelStart = objElement.SelStart - Len(chg$)
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                RTFForm.RichTextBox.SelLength = Len(chg$)
            Else
                RTFForm.RichTextBox.SelStart = RTFForm.RichTextBox.SelStart + Len(chg$)
            End If
        End If
        UndoStack.Add Item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    trapUndo = True
    RTFForm.RichTextBox.SetFocus
End Sub



Public Sub Undo()
Dim chg$, x&
Dim DeleteFlag As Boolean
Dim objElement As Object, objElement2 As Object
    If UndoStack.Count > 1 And trapUndo Then
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then
            x& = SendMessage(RTFForm.RichTextBox.hwnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            RTFForm.RichTextBox.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            RTFForm.RichTextBox.SelLength = objElement.TextLen - objElement2.TextLen
            RTFForm.RichTextBox.SelText = ""
            x& = SendMessage(RTFForm.RichTextBox.hwnd, EM_HIDESELECTION, 0&, 0&)
        Else
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.Text, objElement2.Text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            RTFForm.RichTextBox.SelStart = objElement2.SelStart
            RTFForm.RichTextBox.SelLength = 0
            RTFForm.RichTextBox.SelText = chg$
            RTFForm.RichTextBox.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                RTFForm.RichTextBox.SelLength = Len(chg$)
            Else
                RTFForm.rtfText.SelStart = RTFForm.RichTextBox.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add Item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    trapUndo = True
    RTFForm.RichTextBox.SetFocus
End Sub




Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
Dim tempParam$
Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function

-------------------------------------------


'RTF Form's Code
Private Sub RichTextBox_Change()
    If Not trapUndo Then Exit Sub

    Dim newElement As New UndoElement
    Dim c%, l&

  
    For c% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c%


    newElement.SelStart = RichTextBox.SelStart
    newElement.TextLen = Len(RichTextBox.Text)
    newElement.Text = RichTextBox.Text


    UndoStack.Add Item:=newElement
    
End Sub   


Private Sub Form_Load()
trapUndo = True
RichTextBox_Change
End Sub


Private Sub UndoButton_Click()
Call Undo
End Sub

Private Sub RedoButton_Click()
Call Redo
End Sub
