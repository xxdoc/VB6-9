Function FixNewLineChars(ByVal Text As String) As String

  Dim LineFeed As String * 1
  Dim CarrigeReturn As String * 1
  Dim BeepChar As String * 1

  LineFeed = Chr(10)
  CarrigeReturn = Chr(13)
  BeepChar = Chr(7)
  Text = Replace(Text, vbCrLf, BeepChar)
  Text = Replace(Text, LineFeed, BeepChar)
  Text = Replace(Text, CarrigeReturn, BeepChar)
  Text = Replace(Text, LineFeed & CarrigeReturn, BeepChar)

  FixNewLineChars = Replace(Text, BeepChar, vbCrLf)

End Function
