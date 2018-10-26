Attribute VB_Name = "modLogIt"
Public Function LogIt(TextToLog As String)

    ' ***
    ' *** Function:     LogIt
    ' *** Date:         2018-06-14
    ' *** Author:       fortypoundhead\VB6Boy
    ' *** Purpose:      Create/append log file for the current day, with date/time stamping of each record.
    ' ***
    ' *** Dependencies:
    ' ***               ISO8601DateTime
    ' ***               ISO8601Date
    ' ***
    
    Dim strLogFileName As String
    Dim strLogEntryTime As String
    
    strLogEntryTime = ISO8601DateTime
    strLogFileName = App.Path & "\" & ISO8601Date & ".txt"
    Open strLogFileName For Append As #1
    Print #1, TextToLog
    Close #1

End Function

Public Function ISO8601DateTime() As String

    ' ***
    ' *** Function:     ISO8601DateTime
    ' *** Date:         2018-06-14
    ' *** Author:       fortypoundhead\VB6Boy
    ' *** Purpose:      Return date and time formatted as per ISO-8601 specification, in the format:
    ' ***
    ' ***               YYYY-MM-DD HH:MM:SS
    ' ***
    ' *** Dependencies:
    ' ***               None
    ' ***
    ' *** NOTE:
    ' ***               This function is for string handling demonstration only. The better way to do
    ' ***               this is to use the built-in VB6 function Format$.  Format$ uses the following
    ' ***               syntax to achieve the same result:
    ' ***
    ' ***               ISO8601DateTime = Format$(Now, "yyyy-mm-dd") & " " & Format$(Now,"hh:mm:ss")
    ' ***
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strHour As String
    Dim strMinute As String
    Dim strSecond As String
    
    strYear = Year(Now)
    strMonth = Month(Now)
    strDay = Day(Now)
    strHour = Hour(Now)
    strMinute = Minute(Now)
    strSecond = Second(Now)
    
    If Len(strMonth) < 2 Then strMonth = "0" & strMonth
    If Len(strDay) < 2 Then strDay = "0" & strDay
    If Len(strHour) < 2 Then strHour = "0" & strHour
    If Len(strMinute) < 2 Then strMinute = "0" & strMinute
    If Len(strSecond) < 2 Then strSecond = "0" & strSecond
    
    ISO8601DateTime = strYear & "-" & strMonth & "-" & strDay & " " & strHour & ":" & strMinute & ":" & strSecond

End Function

Public Function ISO8601Date() As String

    ' ***
    ' *** Function:     ISO8601Date
    ' *** Date:         2018-06-14
    ' *** Author:       fortypoundhead\VB6Boy
    ' *** Purpose:      Return date formatted as per ISO-8601 specification, in the format:
    ' ***
    ' ***               YYYY-MM-DD
    ' ***
    ' *** Dependencies:
    ' ***               None
    ' ***
    ' *** NOTE:
    ' ***               This function is for string handling demonstration only. The better way to do
    ' ***               this is to use the built-in VB6 function Format$.  Format$ uses the following
    ' ***               syntax to achieve the same result:
    ' ***
    ' ***               ISO8601Date = Format$(Now, "yyyy-mm-dd")
    ' ***

    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    strYear = Year(Now)
    strMonth = Month(Now)
    strDay = Day(Now)
    
    If Len(strMonth) < 2 Then strMonth = "0" & strMonth
    If Len(strDay) < 2 Then strDay = "0" & strDay
    
    ISO8601Date = strYear & "-" & strMonth & "-" & strDay

End Function
