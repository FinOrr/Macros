' Excel Mailer
' Quickly mail your Excel spreasheet using Oulook
Sub SendEmails()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range
    Set OutApp = CreateObject("Outlook.Application")
    For Each cell In Columns("B").Cells.SpecialCells(xlCellTypeConstants)
        If cell.Value Like "?*@?*.?*" And _
           LCase(Cells(cell.Row, "C").Value) = "yes" Then
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = cell.Value
                .Subject = "Hello"
                .Body = "Dear " & Cells(cell.Row, "A").Value & vbNewLine & vbNewLine & _
                         "Your message text."
                .Send
            End With
        End If
    Next cell
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
