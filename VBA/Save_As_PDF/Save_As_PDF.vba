' Save your excel worksheets as a PDF
Sub SaveAsPDF()
    ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:="C:\YourPath\YourFile.pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
End Sub