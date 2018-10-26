Private Declare Function GetRTTAndHopCount Lib "iphlpapi.dll" (ByVal lDestIPAddr As Long, ByRef lHopCount As Long, ByVal lMaxHops As Long, ByRef lRTT As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Public Function SimplePing(sIPadr As String) As Boolean

    Dim lIPadr      As Long
    Dim lHopsCount  As Long
    Dim lRTT        As Long
    Dim lMaxHops    As Long
    Dim lResult     As Long
    
    Const SUCCESS = 1
    
    lMaxHops = 20               ' should be enough ...
    lIPadr = inet_addr(sIPadr)
    SimplePing = (GetRTTAndHopCount(lIPadr, lHopsCount, lMaxHops, lRTT) = SUCCESS)
    
End Function