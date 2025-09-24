Attribute VB_Name = "Module1"
Sub RefreshAndExportDashboard()
    Dim ws As Worksheet, pt As PivotTable
    Dim exportPath As String, fileName As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Refresh every PivotTable in every worksheet
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
        Next pt
    Next ws

     ' Page setup for Dashboard
    With ThisWorkbook.Worksheets("Dashboard").PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
    End With
    
    ' Export to PDF
    exportPath = ThisWorkbook.Path & "\"
    fileName = "Dashboard_Report_" & Format(Date, "yyyymmdd") & ".pdf"
    
    ThisWorkbook.Sheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=exportPath & fileName, _
        Quality:=xlQualityStandard
    
    MsgBox "Dashboard exported as one-page PDF: " & fileName
End Sub

