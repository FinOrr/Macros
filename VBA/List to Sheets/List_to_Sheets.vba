' Create a new spreadsheet for each unique entry in a list
Sub CreateSheetsFromAList()
    Dim MyCell As Range, MyRange As Range
    Set MyRange = Sheets("Sheet1").Range("A1")
    Set MyRange = Range(MyRange, MyRange.End(xlDown))

    For Each MyCell In MyRange
        Sheets.Add After:=Sheets(Sheets.Count) 'creates a new worksheet
        Sheets(Sheets.Count).Name = MyCell.Value 'names the new worksheet
    Next MyCell
End Sub
