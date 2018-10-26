' #VBIDEUtils#*******************************
' * Programmer Name  : Waty Thierry
' * Web Site : http://www.geocities.com/ResearchTriangle/6311/
' * E-Mail           : waty.thierry@usa.net
' * Date             : 25/11/1999
' * Time             : 15:54

' * Comments  : Convert an integer number to text in French
'**********************************************************

Public Function ConvertirEnText(ValNum As Double) As String

Static Unites(0 To 9) As String
Static Dixaines(0 To 9) As String
Static LesDixaines(0 To 9) As String
Static Milliers(0 To 4) As String

Dim i As Integer
Dim nPosition As Integer
Dim ValNb As Integer
Dim LesZeros As Integer
Dim strResultat As String
Dim strTemp As String
Dim tmpBuff As String

Unites(0) = "zero"
Unites(1) = "un"
Unites(2) = "deux"
Unites(3) = "trois"
Unites(4) = "quatre"
Unites(5) = "cinq"
Unites(6) = "six"
Unites(7) = "sept"
Unites(8) = "huit"
Unites(9) = "neuf"

Dixaines(0) = "dix"
Dixaines(1) = "onze"
Dixaines(2) = "douze"
Dixaines(3) = "treize"
Dixaines(4) = "quatorze"
Dixaines(5) = "quinze"
Dixaines(6) = "seize"
Dixaines(7) = "dix-sept"
Dixaines(8) = "dix-huit"
Dixaines(9) = "dix-neuf"

LesDixaines(0) = ""
LesDixaines(1) = "dix"
LesDixaines(2) = "vingt"
LesDixaines(3) = "trente"
LesDixaines(4) = "quarante"
LesDixaines(5) = "cinquante"
LesDixaines(6) = "soixante"
LesDixaines(7) = "soixante-dix"
LesDixaines(8) = "quatre-vingt"
LesDixaines(9) = "quatre-vingt-dix"

Milliers(0) = ""
Milliers(1) = "mille"
Milliers(2) = "million"
Milliers(3) = "millard"
Milliers(4) = "mille"

On Error GoTo NbVersTexteError

strTemp = CStr(Int(ValNum))

For i = Len(strTemp) To 1 Step -1
   ValNb = Val(Mid$(strTemp, i, 1))
   nPosition = (Len(strTemp) - i) + 1
   Select Case (nPosition Mod 3)
      Case 1
         LesZeros = False
         If i = 1 Then
            If ValNb > 1 Then
               tmpBuff = Unites(ValNb) & " "
            Else
               tmpBuff = ""
            End If
         ElseIf Mid$(strTemp, i - 1, 1) = "1" Then
            tmpBuff = Dixaines(ValNb) & " "
            i = i - 1
         ElseIf Mid$(strTemp, i - 1, 1) = "9" Then
            tmpBuff = LesDixaines(8) & " " & _
                Dixaines(ValNb) & " "
            i = i - 1
         ElseIf Mid$(strTemp, i - 1, 1) = "7" Then
            tmpBuff = LesDixaines(6) & " " & _
                 Dixaines(ValNb) & " "
            i = i - 1
         ElseIf ValNb > 0 Then
            tmpBuff = Unites(ValNb) & " "
         Else
            LesZeros = True
            If i > 1 Then
               If Mid$(strTemp, i - 1, 1) <> "0" Then
                  LesZeros = False
               End If
            End If
            If i > 2 Then
               If Mid$(strTemp, i - 2, 1) <> "0" Then
                  LesZeros = False
               End If
            End If
            tmpBuff = ""
         End If
         If LesZeros = False And nPosition > 1 Then
            tmpBuff = tmpBuff & Milliers(nPosition / 3) & " "
         End If
         strResultat = tmpBuff & strResultat
      Case 2
         If ValNb > 0 Then
            strResultat = LesDixaines(ValNb) & " " & _
               strResultat
         End If
      Case 0
         If ValNb > 0 Then
            If ValNb > 1 Then
               strResultat = Unites(ValNb) & " cent " & _
                  strResultat
            Else
               strResultat = "cent " & strResultat
            End If
         End If
   End Select
Next i
If Len(strResultat) > 0 Then
   strResultat = UCase$(Left$(strResultat, 1)) & _
      Mid$(strResultat, 2)
End If

EndNbVersTexte:
ConvertirEnText = strResultat
Exit Function

NbVersTexteError:
strResultat = "Une Erreur !"
Resume EndNbVersTexte
End Function