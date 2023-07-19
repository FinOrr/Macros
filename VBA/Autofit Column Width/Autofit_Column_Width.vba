' No more messing around, let the macro resize your columns to be readable
Sub AutoFitAllColumns()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.EntireColumn.AutoFit
    Next ws
End Sub
