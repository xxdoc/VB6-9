While InStr(1, Text1, "  ") > 0
    Text1 = Replace(Text1, "  ", " ")
Wend

wordCount = Split(Text1, " ")
NumOfWords = UBound(wordCount) + 1
