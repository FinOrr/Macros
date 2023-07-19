Sub DeleteEmptyRows()
    On Error Resume Next
    With ThisWorkbook.ActiveSheet
        .Cells.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    End With
    On Error GoTo 0
End Sub
