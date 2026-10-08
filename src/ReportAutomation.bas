Attribute VB_Name = "ReportAutomation"
Option Explicit

' ReportAutomation
' ------------------
' Consolidates Financial_report_module.bas (Module1, first draft) and
' finautomation_templateexceltoword.bas (Module2, improved draft) into one
' module that compiles and runs repeatedly without manual cleanup between
' runs. See legacy/ for the originals and the defects each had.
'
' Expected workbook layout:
'   DataSources       Column A, from row 2 down: full paths to source workbooks.
'   RawData           Cleared and refilled on every run. Columns A-F copied from
'                      each source workbook's first sheet: Region, Month, Revenue,
'                      Expenses, CustomerName, CompanyName.
'   CalculatedMetrics Columns A-D, one row per RawData row: Net Profit,
'                      Revenue Growth (%), Profit Margin (%), Total Revenue.
'   TemplateInfo      B1 = path to the Word contract template (.docx).
'                      B2 = output folder for generated contracts; created if
'                      it does not already exist.
'
' Bookmarks expected in the Word template: RegionBookmark, MonthBookmark,
' RevenueBookmark, ExpensesBookmark, NetProfitBookmark, CustomerNameBookmark,
' CompanyNameBookmark.

Sub AutomateFinancialReporting()
    Dim wsSources As Worksheet
    Dim wsData As Worksheet
    Dim wsMetrics As Worksheet
    Dim wb As Workbook
    Dim sourceSheet As Worksheet
    Dim sourcePath As String
    Dim sourceRow As Long
    Dim destRow As Long
    Dim lastSourceDataRow As Long

    On Error GoTo ErrorHandler

    Set wsSources = ThisWorkbook.Sheets("DataSources")
    Set wsData = ThisWorkbook.Sheets("RawData")
    Set wsMetrics = ThisWorkbook.Sheets("CalculatedMetrics")

    wsData.Cells.ClearContents

    ' sourceRow walks the DataSources list; destRow walks RawData. The original
    ' code read the next DataSources path with the RawData row counter, so once a
    ' source workbook contributed more or fewer than one row the two sheets fell
    ' out of step and later paths were skipped or read twice. Keeping them as two
    ' independent counters avoids that.
    sourceRow = 2
    destRow = 2

    sourcePath = wsSources.Cells(sourceRow, 1).Value
    Do While sourcePath <> ""
        If Dir(sourcePath) = "" Then
            MsgBox "File not found: " & sourcePath, vbExclamation
            GoTo NextSource
        End If

        Set wb = Workbooks.Open(sourcePath)
        Set sourceSheet = wb.Sheets(1)
        lastSourceDataRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row

        If lastSourceDataRow < 2 Then
            MsgBox "No data found in: " & sourcePath, vbExclamation
            wb.Close False
            Set wb = Nothing
            GoTo NextSource
        End If

        sourceSheet.Range("A2:F" & lastSourceDataRow).Copy wsData.Cells(destRow, 1)
        destRow = destRow + (lastSourceDataRow - 1)

        wb.Close False
        Set wb = Nothing

NextSource:
        sourceRow = sourceRow + 1
        sourcePath = wsSources.Cells(sourceRow, 1).Value
    Loop

    CalculateMetrics wsData, wsMetrics
    FormatSummaryReport wsMetrics
    CreateCharts wsMetrics

    MsgBox "Monthly Financial Report automation complete!"
    Exit Sub

ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    MsgBox "An error occurred in AutomateFinancialReporting: " & Err.Description, vbCritical
End Sub

Sub CalculateMetrics(wsData As Worksheet, wsMetrics As Worksheet)
    Dim lastRow As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    wsMetrics.Cells.ClearContents
    wsMetrics.Range("A1").Value = "Net Profit"
    wsMetrics.Range("B1").Value = "Revenue Growth (%)"
    wsMetrics.Range("C1").Value = "Profit Margin (%)"
    wsMetrics.Range("D1").Value = "Total Revenue"

    For i = 2 To lastRow
        ' Net Profit
        wsMetrics.Cells(i, 1).Value = wsData.Cells(i, 3).Value - wsData.Cells(i, 4).Value

        ' Revenue Growth (%) - assumes data is ordered by month. Guards the
        ' previous month's revenue being zero, which otherwise divides by zero.
        If i > 2 And wsData.Cells(i - 1, 3).Value <> 0 Then
            wsMetrics.Cells(i, 2).Value = (wsData.Cells(i, 3).Value - wsData.Cells(i - 1, 3).Value) / wsData.Cells(i - 1, 3).Value * 100
        Else
            wsMetrics.Cells(i, 2).Value = "N/A"
        End If

        ' Profit Margin (%)
        If wsData.Cells(i, 3).Value <> 0 Then
            wsMetrics.Cells(i, 3).Value = (wsMetrics.Cells(i, 1).Value / wsData.Cells(i, 3).Value) * 100
        Else
            wsMetrics.Cells(i, 3).Value = "N/A"
        End If

        ' Total Revenue
        wsMetrics.Cells(i, 4).Value = wsData.Cells(i, 3).Value
    Next i
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in CalculateMetrics: " & Err.Description, vbCritical
End Sub

Sub FormatSummaryReport(wsMetrics As Worksheet)
    Dim lastRow As Long
    Dim lo As ListObject

    On Error GoTo ErrorHandler

    lastRow = wsMetrics.Cells(wsMetrics.Rows.Count, "A").End(xlUp).Row

    With wsMetrics
        .Range("A1:D1").Font.Bold = True
        .Columns("A:D").AutoFit

        ' ListObjects.Add fails with "a table already exists" on a second run.
        ' Unlist removes the table formatting (not the data) so Add can recreate
        ' it cleanly every time.
        For Each lo In .ListObjects
            lo.Unlist
        Next lo

        .ListObjects.Add(xlSrcRange, .Range("A1:D" & lastRow), , xlYes).TableStyle = "TableStyleMedium9"
    End With
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in FormatSummaryReport: " & Err.Description, vbCritical
End Sub

Sub CreateCharts(wsMetrics As Worksheet)
    Dim lastRow As Long
    Dim chartObj As ChartObject

    On Error GoTo ErrorHandler

    lastRow = wsMetrics.Cells(wsMetrics.Rows.Count, "A").End(xlUp).Row

    ' Remove a chart left over from a previous run instead of stacking a new one
    ' on top of it.
    For Each chartObj In wsMetrics.ChartObjects
        chartObj.Delete
    Next chartObj

    Set chartObj = wsMetrics.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        ' PlotBy:=xlColumns is explicit on purpose. Both columns hold numbers, so
        ' Excel's own guess at rows-vs-columns is unreliable: left to itself it
        ' picked the first data row as a series name ("Net Profit $40,000.00" in
        ' the legend) and the hardcoded A1:B10 range only ever covered 9 of the
        ' 12 months. Reading lastRow and naming the series from the header cells
        ' fixes both.
        .SetSourceData Source:=wsMetrics.Range("A1:B" & lastRow), PlotBy:=xlColumns
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Monthly Net Profit and Revenue Growth"
        .SeriesCollection(1).Name = "=""" & wsMetrics.Range("A1").Value & """"
        .SeriesCollection(2).Name = "=""" & wsMetrics.Range("B1").Value & """"
        ' Growth is a percentage in the single digits; on the dollar axis it
        ' renders as invisible bars at zero, so it gets its own axis as a line.
        .SeriesCollection(2).ChartType = xlLine
        .SeriesCollection(2).AxisGroup = xlSecondary
    End With
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in CreateCharts: " & Err.Description, vbCritical
End Sub

Sub TransferToWordTemplate()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templateSheet As Worksheet
    Dim templatePath As String
    Dim outputFolder As String
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim region As String, month As String
    Dim revenue As Double, expenses As Double, netProfit As Double
    Dim customerName As String, companyName As String
    Dim outputPath As String
    Dim missing As String

    On Error GoTo ErrorHandler

    Set templateSheet = ThisWorkbook.Sheets("TemplateInfo")
    templatePath = templateSheet.Range("B1").Value
    outputFolder = templateSheet.Range("B2").Value
    If Right(outputFolder, 1) = "\" Then outputFolder = Left(outputFolder, Len(outputFolder) - 1)

    If Dir(templatePath) = "" Then
        MsgBox "The Word template file could not be found. Check TemplateInfo!B1.", vbExclamation
        Exit Sub
    End If

    EnsureFolderExists outputFolder

    Set wsData = ThisWorkbook.Sheets("RawData")
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    For i = 2 To lastRow
        If IsEmpty(wsData.Cells(i, 1)) Or IsEmpty(wsData.Cells(i, 2)) Or _
           IsEmpty(wsData.Cells(i, 3)) Or IsEmpty(wsData.Cells(i, 4)) Or _
           IsEmpty(wsData.Cells(i, 5)) Or IsEmpty(wsData.Cells(i, 6)) Then
            MsgBox "Missing data in row " & i & ". Skipping.", vbExclamation
            GoTo NextDataRow
        End If

        region = wsData.Cells(i, 1).Value
        month = wsData.Cells(i, 2).Value
        revenue = wsData.Cells(i, 3).Value
        expenses = wsData.Cells(i, 4).Value
        netProfit = revenue - expenses
        customerName = wsData.Cells(i, 5).Value
        companyName = wsData.Cells(i, 6).Value

        ' A fresh document per row, instead of reusing one wordDoc across the
        ' whole loop. A bookmark's range collapses to its filled text the moment
        ' it is written, so re-using the same document meant only the first row
        ' ever found its bookmarks; every later row silently filled nothing.
        Set wordDoc = wordApp.Documents.Add(Template:=templatePath, NewTemplate:=False)

        FillBookmark wordDoc, "RegionBookmark", region, i, missing
        FillBookmark wordDoc, "MonthBookmark", month, i, missing
        FillBookmark wordDoc, "RevenueBookmark", Format(revenue, "$#,##0.00"), i, missing
        FillBookmark wordDoc, "ExpensesBookmark", Format(expenses, "$#,##0.00"), i, missing
        FillBookmark wordDoc, "NetProfitBookmark", Format(netProfit, "$#,##0.00"), i, missing
        FillBookmark wordDoc, "CustomerNameBookmark", customerName, i, missing
        FillBookmark wordDoc, "CompanyNameBookmark", companyName, i, missing

        outputPath = outputFolder & "\" & region & "_" & month & "_Contract.docx"
        wordDoc.SaveAs2 outputPath, FileFormat:=16 ' wdFormatXMLDocument, i.e. .docx

        wordDoc.Close SaveChanges:=False
        Set wordDoc = Nothing

NextDataRow:
    Next i

    wordApp.Quit
    Set wordApp = Nothing

    If Len(missing) > 0 Then
        MsgBox "Contracts written to " & outputFolder & vbCrLf & vbCrLf & _
               "Bookmarks missing from the template (left unfilled):" & missing, vbExclamation
    Else
        MsgBox "Contracts written to " & outputFolder, vbInformation
    End If
    Exit Sub

ErrorHandler:
    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
        Set wordDoc = Nothing
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit
        Set wordApp = Nothing
    End If
    MsgBox "An error occurred in TransferToWordTemplate: " & Err.Description, vbCritical
End Sub

Private Sub FillBookmark(wordDoc As Object, ByVal name As String, ByVal text As String, _
                         ByVal rowNum As Long, ByRef missing As String)
    If wordDoc.Bookmarks.Exists(name) Then
        wordDoc.Bookmarks.Item(name).Range.Text = text
    Else
        missing = missing & vbCrLf & "  row " & rowNum & ": " & name
    End If
End Sub

Private Sub EnsureFolderExists(ByVal folderPath As String)
    ' Creates folderPath one level at a time (MkDir fails if the immediate
    ' parent does not exist yet, so a straight MkDir on a multi-level path
    ' that is missing more than its last segment would fail).
    Dim parts() As String
    Dim currentPath As String
    Dim idx As Long

    If folderPath = "" Then Exit Sub
    If Dir(folderPath, vbDirectory) <> "" Then Exit Sub

    parts = Split(folderPath, "\")
    currentPath = parts(0)

    For idx = 1 To UBound(parts)
        currentPath = currentPath & "\" & parts(idx)
        If Dir(currentPath, vbDirectory) = "" Then MkDir currentPath
    Next idx
End Sub
