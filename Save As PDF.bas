Attribute VB_Name = "Module1"
Sub SavePDF()

Application.PrintCommunication = False

    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        '.PrintArea = Worksheets(ReportWsName).UsedRange
        .FitToPagesWide = 1
        '.FitToPagesTall = 1
    End With
    
    Dim relativePath As String
    relativePath = ThisWorkbook.Path & "\" & "Export.pdf"

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=relativePath, _
OpenAfterPublish:=False

Application.PrintCommunication = True

End Sub

