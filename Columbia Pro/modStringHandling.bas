Attribute VB_Name = "modStringHandling"
Option Explicit

Const MAXCHARS = 10000      ' maximum number of characters to post in a scrolling textbox

Public Function ParameterCount(ParseCharacter As String, tString As Variant) As Integer

    '***************************************************************************************
    '***
    '*** FUNCTION: ParameterCount
    '*** AUTHOR: Derek A. Wirch
    '*** DATE: 09/26/2002
    '*** COPYRIGHT: (c)2003 Derek A. Wirch
    '***
    '*** PURPOSE: Counts the number of fields in a delimited string, given the
    '*** the delimiter and the string
    '*** RETURNS: Number of fields in the string
    '*** USAGE: ret = ParameterCount(",",txtString)
    '***
    '*** SIDE EFFECTS: None
    '***
    '***************************************************************************************
    '***
    '*** HISTORY:
    '*** Date Description
    '*** ---------- --------------------------------------------------------------------
    '*** 07/30/2003 Initial Creation
    '*** 09/25/2003 Addition to main string handling module
    '***
    '***************************************************************************************
    
    Dim CurrentPosition As Integer
    Dim ParseToPosition As Integer
    Dim CurrentToken As Integer
    Dim tempstring As String
    
    tempstring = Trim(tString) + ParseCharacter
    
    If Len(tempstring) = 1 Then Exit Function
    
    CurrentPosition = 1
    CurrentToken = 1
    
    Do
        ParseToPosition = InStr(CurrentPosition, tempstring, ParseCharacter)
        CurrentToken = CurrentToken + 1
        CurrentPosition = ParseToPosition + 1
    
    Loop Until (CurrentPosition >= Len(tempstring))
    
    ParameterCount = CurrentToken - 1

End Function

Public Function ParameterValue(ParseCharacter As String, tString As Variant, Index As Integer) As String

    '***************************************************************************************
    '***
    '*** FUNCTION: ParameterValue
    '*** AUTHOR: Derek A. Wirch
    '*** DATE: 09/26/2002
    '*** COPYRIGHT: (c)2003 Derek A. Wirch
    '***
    '*** PURPOSE: Returns a field value from a delimited string, given the delimiter,
    '*** string, and which field to pick
    '*** RETURNS: string value of field in the string
    '*** USAGE: outstring = ParameterValue(",",instring,3)
    '***
    '*** SIDE EFFECTS: None
    '***
    '***************************************************************************************
    '***
    '*** HISTORY:
    '*** Date Description
    '*** ---------- --------------------------------------------------------------------
    '*** 07/30/2003 Initial Creation
    '*** 09/25/2003 Addition to main string handling module
    '***
    '***************************************************************************************
    
    Dim CurrentPosition As Integer
    Dim ParseToPosition As Integer
    Dim CurrentToken As Integer
    Dim tempstring As String
    
    tempstring = Trim(tString) + ParseCharacter
    
    If Len(tempstring) = 1 Then Exit Function
    
    CurrentPosition = 1
    CurrentToken = 1
    
    Do
        ParseToPosition = InStr(CurrentPosition, tempstring, ParseCharacter)
        
        If Index = CurrentToken Then
        
            ParameterValue = Mid$(tempstring, CurrentPosition, ParseToPosition - CurrentPosition)
            Exit Function
        
        End If
        
        CurrentToken = CurrentToken + 1
        CurrentPosition = ParseToPosition + 1
    
    Loop Until (CurrentPosition >= Len(tempstring))

End Function

Public Sub Post(tbxEditBox As TextBox, sNewText As String, Optional SendToLog As Boolean = False)

    '----------------------------------------------------------------------------
    ' Procedure : Post
    ' Author    : Derek A. Wirch
    ' Purpose   : Makes a multi-line textbox into a scrolling textbox, chat-style.
    ' Returns   : N/A
    ' Syntax    : Post txtTextBox, "Text to Post"
    ' Side Eff. : N/A
    '
    ' Modification History:
    '
    ' Date       Mod Notes
    ' ---------- ----------------------------------------------------------------
    ' 2006/08/31 Initial Creation
    '
    '----------------------------------------------------------------------------
    
    Dim LogText As String
    
    LogText = sNewText
    sNewText = sNewText & vbCrLf
    
    With tbxEditBox
        If Len(sNewText) + Len(.Text) > MAXCHARS Then
            'Scroll some text off the top to make more room
            .Text = Mid$(.Text, InStr(100 + Len(sNewText), .Text, vbCrLf) + 2)
        End If
        .SelStart = Len(.Text)
        .SelText = sNewText
    End With
    DoEvents
    If SendToLog = True Then
        SendToLogFile LogText
    End If
    
End Sub
Public Function SendToLogFile(StringToWrite As String)

    Open LogDir & uDate & ".log" For Append As #1
    Print #1, uDateTime & " " & StringToWrite
    Close #1
    
End Function
Public Function uDateTime() As String

    '***
    '*** Create date/time string in the format YYYY-MM-DD HH:MM:SS
    '***
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strHours As String
    Dim strMinutes As String
    Dim strSeconds As String
    
    strYear = Trim(Year(Now))
    strMonth = Trim(Month(Now))
    strDay = Trim(Day(Now))
    strHours = Trim(Hour(Now))
    strMinutes = Trim(Minute(Now))
    strSeconds = Trim(Second(Now))
    
    If Len(strMonth) < 2 Then strMonth = "0" & strMonth
    If Len(strDay) < 2 Then strDay = "0" & strDay
    If Len(strHours) < 2 Then strHours = "0" & strHours
    If Len(strSeconds) < 2 Then strSeconds = "0" & strSeconds
    
    uDateTime = strYear & "-" & strMonth & "-" & strDay & " " & strHours & ":" & strMinutes & ":" & strSeconds
    
End Function

Public Function uDate() As String

    '***
    '*** Create date string in the format YYYY-MM-DD
    '***
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    strYear = Trim(Year(Now))
    strMonth = Trim(Month(Now))
    strDay = Trim(Day(Now))
    
    If Len(strMonth) < 2 Then strMonth = "0" & strMonth
    If Len(strDay) < 2 Then strDay = "0" & strDay
    
    uDate = strYear & "-" & strMonth & "-" & strDay
    
End Function
