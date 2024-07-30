# Daily_TMS_Chart_Generation
Generates charts for each TMS based on BUNO count of aircraft inventory status
Introduction
This guide will walk you through the process of using a VBA script to create new tabs for each unique value in a column, generate pivot tables, and create stacked column charts with custom colors in Excel. The script is designed to work with data that includes columns for TMS, Acft Inv Term, BUNO, and Service.
Prerequisites
- Basic understanding of Excel and VBA.
- Excel file with columns named "TMS" (column B), "Acft Inv Term" (column AF), "BUNO" (column D), and "Service" (column AG).
Steps
1. Open Excel and Prepare Your Data
Ensure your data is structured correctly with the appropriate column headers. Your worksheet should look something like this:

| TMS  | ... | BUNO | ... | Acft Inv Term | Service |
|------|-----|------|-----|---------------|---------|
| Data | ... | Data | ... | Data          | Data    |
2. Open the VBA Editor
1. Press `Alt + F11` to open the VBA editor.
3. Insert a New Module
1. In the VBA editor, click `Insert > Module` to create a new module.
2. Copy and paste the following VBA code into the new module:

Sub CreateTabsAndCharts()
    Dim ws As Worksheet
    Dim uniqueTMS As Collection
    Dim cell As Range
    Dim TMS As Variant
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim chartObj As ChartObject
    Dim tmsDict As Object
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim dataCount As Long
    Dim pf As PivotField

    ' Initialize the collection and dictionary
    Set uniqueTMS = New Collection
    Set tmsDict = CreateObject("Scripting.Dictionary")
    
    ' Set the source worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name

    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Collect unique TMS values
    On Error Resume Next
    For Each cell In ws.Range("B2:B" & lastRow)
        If Not tmsDict.exists(cell.Value) Then
            uniqueTMS.Add cell.Value, CStr(cell.Value)
            tmsDict.Add cell.Value, 1
        End If
    Next cell
    On Error GoTo 0

    ' Loop through each unique TMS value to create a new sheet and chart
    For Each TMS In uniqueTMS
        ' Add a new sheet and name it
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = TMS

        ' Filter the data for the current TMS value
        ws.Rows(1).AutoFilter Field:=2, Criteria1:=TMS
        ws.Range("A1:AG" & lastRow).SpecialCells(xlCellTypeVisible).Copy Destination:=newSheet.Range("A1")

        ' Clear any filters
        ws.AutoFilterMode = False

        ' Check if there is more than one row of data
        dataCount = newSheet.Cells(newSheet.Rows.Count, 1).End(xlUp).Row
        If dataCount > 1 Then
            ' Create a pivot table to count BUNO by Acft Inv Term and Service
            Set dataRange = newSheet.Range("A1:AG" & dataCount)
            Set pivotCache = ThisWorkbook.PivotCaches.Create(xlDatabase, dataRange)
            Set pivotTable = newSheet.PivotTables.Add(PivotCache:=pivotCache, TableDestination:=newSheet.Range("AH1"), TableName:="PivotTable_" & TMS)
            
            With pivotTable
                .PivotFields("Acft Inv Term").Orientation = xlRowField
                .PivotFields("Service").Orientation = xlColumnField
            End With
            
            ' Add "BUNO" to the data fields and set it to count
            Set pf = pivotTable.PivotFields("BUNO")
            pf.Orientation = xlDataField
            pf.Function = xlCount
            pf.Name = "Count of BUNO"
            
            ' Create a chart from the pivot table
            Set chartObj = newSheet.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
            With chartObj.Chart
                .SetSourceData Source:=pivotTable.TableRange2
                .ChartType = xlColumnStacked
                .HasTitle = True
                .ChartTitle.Text = "Count of BUNO by Acft Inv Term for " & TMS
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Acft Inv Term"
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "Count of BUNO"
                .HasLegend = True

                ' Hide all field buttons
                .ShowAllFieldButtons = False

                ' Add data labels
                .ApplyDataLabels
            End With
            
            ' Customize colors for "Service" values
            Dim series As Series
            For Each series In chartObj.Chart.SeriesCollection
                Select Case series.Name
                    Case "USMC"
                        series.Format.Fill.ForeColor.RGB = RGB(0, 128, 0) ' Green
                    Case "USN"
                        series.Format.Fill.ForeColor.RGB = RGB(0, 0, 255) ' Blue
                End Select
            Next series
        Else
            MsgBox "Not enough data for TMS: " & TMS, vbExclamation
        End If
    Next TMS
End Sub


Sub CreateTabsAndCharts()
    Dim ws As Worksheet
    Dim uniqueTMS As Collection
    Dim cell As Range
    Dim TMS As Variant
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim chartObj As ChartObject
    Dim tmsDict As Object
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim dataCount As Long
    Dim pf As PivotField

    ' Initialize the collection and dictionary
    Set uniqueTMS = New Collection
    Set tmsDict = CreateObject("Scripting.Dictionary")
    
    ' Set the source worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name

    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Collect unique TMS values
    On Error Resume Next
    For Each cell In ws.Range("B2:B" & lastRow)
        If Not tmsDict.exists(cell.Value) Then
            uniqueTMS.Add cell.Value, CStr(cell.Value)
            tmsDict.Add cell.Value, 1
        End If
    Next cell
    On Error GoTo 0

    ' Loop through each unique TMS value to create a new sheet and chart
    For Each TMS In uniqueTMS
        ' Add a new sheet and name it
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = TMS

        ' Filter the data for the current TMS value
        ws.Rows(1).AutoFilter Field:=2, Criteria1:=TMS
        ws.Range("A1:AG" & lastRow).SpecialCells(xlCellTypeVisible).Copy Destination:=newSheet.Range("A1")

        ' Clear any filters
        ws.AutoFilterMode = False

        ' Check if there is more than one row of data
        dataCount = newSheet.Cells(newSheet.Rows.Count, 1).End(xlUp).Row
        If dataCount > 1 Then
            ' Create a pivot table to count BUNO by Acft Inv Term and Service
            Set dataRange = newSheet.Range("A1:AG" & dataCount)
            Set pivotCache = ThisWorkbook.PivotCaches.Create(xlDatabase, dataRange)
            Set pivotTable = newSheet.PivotTables.Add(PivotCache:=pivotCache, TableDestination:=newSheet.Range("AH1"), TableName:="PivotTable_" & TMS)
            
            With pivotTable
                .PivotFields("Acft Inv Term").Orientation = xlRowField
                .PivotFields("Service").Orientation = xlColumnField
            End With
            
            ' Add "BUNO" to the data fields and set it to count
            Set pf = pivotTable.PivotFields("BUNO")
            pf.Orientation = xlDataField
            pf.Function = xlCount
            pf.Name = "Count of BUNO"
            
            ' Create a chart from the pivot table
            Set chartObj = newSheet.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
            With chartObj.Chart
                .SetSourceData Source:=pivotTable.TableRange2
                .ChartType = xlColumnStacked
                .HasTitle = True
                .ChartTitle.Text = "Count of BUNO by Acft Inv Term for " & TMS
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Acft Inv Term"
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "Count of BUNO"
                .HasLegend = True

                ' Hide all field buttons
                .ShowAllFieldButtons = False

                ' Add data labels
                .ApplyDataLabels
            End With
            
            ' Customize colors for "Service" values
            Dim series As Series
            For Each series In chartObj.Chart.SeriesCollection
                  Select Case series.Name
                    Case "USMC"
                        series.Format.Fill.ForeColor.RGB = RGB(0, 128, 0) ' Green
                    Case "USN"
                        series.Format.Fill.ForeColor.RGB = RGB(173, 216, 230) ' Light Blue
                    Case "CNATRA"
                        series.Format.Fill.ForeColor.RGB = RGB(139, 69, 19) ' Brown
                    Case "NAVAIR"
                        series.Format.Fill.ForeColor.RGB = RGB(255, 204, 0) ' Dark Yellow
                    Case "Misc"
                        series.Format.Fill.ForeColor.RGB = RGB(169, 169, 169) ' Grey
                End Select
            Next series
        Else
            MsgBox "Not enough data for TMS: " & TMS, vbExclamation
        End If
    Next TMS
End Sub

4. Run the Macro
1. Close the VBA editor.
2. Press `Alt + F8` to open the Macro dialog box.
3. Select `CreateTabsAndCharts` and click `Run`.
5. Review the Results
The macro will create new tabs for each unique TMS value, generate pivot tables, and create stacked column charts with custom colors for "USMC" (green) and "USN" (blue).
Troubleshooting
- Ensure that your data is correctly formatted and that the column headers match the expected names ("TMS", "Acft Inv Term", "BUNO", and "Service").
- If the macro encounters an error, double-check the data structure and ensure that there are no empty rows or columns within the data range.
Conclusion
This guide has walked you through the process of using a VBA script to automate the creation of tabs and charts in Excel. By following these steps, you can efficiently analyze and visualize your data based on unique values in the "TMS" column.
