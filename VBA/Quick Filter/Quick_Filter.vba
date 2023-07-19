' Quick filer based on known criteria
Sub FilterData()
    Dim rng As Range
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set rng = ws.Range("A1:D100")

    rng.AutoFilter Field:=1, Criteria1:="YourCriteria"
End Sub
