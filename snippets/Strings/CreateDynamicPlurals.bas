'Syntax     IIF(Expression,true,false)

'Samples:

'This sample prints an "s" after the string in Array element 
'FType(0,1) if the
'value in element FType(1,1) is does not equal than 1

Document.print IIf(FType(1, 1) <> 1, FType(0, 1) & "s", _
  FType(0, 1))

'This sample adds the s to a datafield

IIf(MyDB.fields("Quantity") <> 1, MyDB.fields("Item")= ItemName _
& "s", MyDB.fields("Item")= ItemName)