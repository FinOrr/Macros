' Combine data from multiple sheets into a single one
Sub ConsolidateWorksheets()
    Dim ws As Worksheet
    Dim wsMerged As Worksheet

    Set wsMerged = ThisWorkbook.Sheets.Add
    wsMerged.Name = "Merged"

    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Merged" Then
            ws.UsedRange.Copy wsMerged.Cells(wsMerged.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
        End If
    Next ws
End Sub
