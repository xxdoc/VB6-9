Private Function Char(Num As Long, Optional Special As Integer = 256) As String
 
    ' ***
    ' *** Special must not be higher then 256 or lower then 0
    ' ***

    If Special < 1 Then
        Num = 1
    ElseIf Special > 256 Then
        Num = 256
    End If
    
    ' ***
    ' *** If passed Num is higher then 255, it keeps subtracting
    ' *** the Num by Special(256) until Num becomes legal ASCII number.
    ' ***     
    
    If Num > 255 Then
        While Num > 255
            Num = Num - Special
        Wend
    End If
    
    ' *** 
    ' *** If Num is lower then 0, it keeps adding Special(256)
    ' *** until Num becomes legal ASCII number.
    ' *** 
    
    If Num < 0 Then
        While Num < 0
            Num = Num + Special
        Wend
    End If
    
    ' *** 
    ' *** At the end it just passes the Num to Chr() function
    ' *** 
    
    Char = Chr(Num)
 
End Function