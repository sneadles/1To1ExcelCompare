Attribute VB_Name = "Module11"
Sub SpecialCharacters()

maxrows = Range("A" & Rows.Count).End(xlUp).Row
maxcolumns = Cells(1, Columns.Count).End(xlToLeft).Column

For i = 1 To maxrows
If Sheet1.Cells(i, 1).Value <> Sheet2.Cells(i, 1).Value Then Sheet1.Cells(i, 1).Interior.ColorIndex = 3

b = 1
Do While b < maxcolumns
If Sheet1.Cells(i, b).Value <> Sheet2.Cells(i, b).Value Then Sheet1.Cells(i, b).Interior.ColorIndex = 3
b = b + 1

Loop
Next i

End Sub
