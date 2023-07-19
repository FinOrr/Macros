' If you've got pivot tables that need constant refreshing, check it out
Sub RefreshAllPivotTables()
    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub
