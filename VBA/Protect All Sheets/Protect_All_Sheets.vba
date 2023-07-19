' One click password locking for all sheets in a file
' If you want to send out your spreadsheet as read-only
Sub ProtectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:="YourPassword"
    Next ws
End Sub
