Option Explicit

'================================================================================================
' Product Matching Engine - Excel Threshold Explorer Automation Macro
'================================================================================================
' This macro automates the setup of threshold analysis in Excel using the enhanced export data
' from the Product Matching Engine.
'
' Requirements:
' - Excel with Power Pivot enabled (Excel 2016+ or Excel 2013 with Power Pivot add-in)
' - Enhanced threshold analysis export from the Product Matching Engine
'
' Usage:
' 1. Open the enhanced_threshold_analysis.xlsx file
' 2. Press Alt+F11 to open VBA editor
' 3. Insert this code in a new module
' 4. Run the "SetupThresholdExplorer" macro
'================================================================================================

Sub SetupThresholdExplorer()
    ' Main subroutine to set up the complete threshold explorer
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Step 1: Add data to Power Pivot model
    If Not AddToDataModel(wb) Then
        MsgBox "Failed to add tables to Power Pivot data model", vbExclamation
        Exit Sub
    End If
    
    ' Step 2: Create relationships
    If Not CreateRelationships() Then
        MsgBox "Failed to create relationships", vbExclamation
        Exit Sub
    End If
    
    ' Step 3: Add calculated columns
    AddCalculatedColumns
    
    ' Step 4: Add measures
    AddMeasures
    
    ' Step 5: Create analysis sheet
    CreateAnalysisSheet
    
    ' Step 6: Create dashboard
    CreateDashboard
    
    ' Step 7: Add instructions
    AddInstructions
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    MsgBox "Threshold Explorer setup complete!" & vbCrLf & vbCrLf & _
           "Sheets created:" & vbCrLf & _
           "- Analysis (for custom threshold testing)" & vbCrLf & _
           "- Dashboard (with charts and slicers)" & vbCrLf & _
           "- Instructions (quick reference)", _
           vbInformation, "Setup Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Function AddToDataModel(wb As Workbook) As Boolean
    ' Add tables to Power Pivot data model
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tableName As String
    
    ' Add Product List table
    Set ws = wb.Sheets("Product List")
    If ws Is Nothing Then
        MsgBox "Product List sheet not found!", vbExclamation
        GoTo ErrorHandler
    End If
    
    tableName = "ProductList"
    ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = tableName
    ActiveWorkbook.Model.ModelTables.Add tableName, ws.ListObjects(1)
    
    ' Add Similarity Matrix table
    Set ws = wb.Sheets("Similarity Matrix")
    If ws Is Nothing Then
        MsgBox "Similarity Matrix sheet not found!", vbExclamation
        GoTo ErrorHandler
    End If
    
    tableName = "SimilarityMatrix"
    ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = tableName
    ActiveWorkbook.Model.ModelTables.Add tableName, ws.ListObjects(1)
    
    ' Add Threshold Summary table
    Set ws = wb.Sheets("Threshold Summary")
    If Not ws Is Nothing Then
        tableName = "ThresholdSummary"
        ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = tableName
        ActiveWorkbook.Model.ModelTables.Add tableName, ws.ListObjects(1)
    End If
    
    ' Add Threshold Groups table
    Set ws = wb.Sheets("Threshold Groups")
    If Not ws Is Nothing Then
        tableName = "ThresholdGroups"
        ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = tableName
        ActiveWorkbook.Model.ModelTables.Add tableName, ws.ListObjects(1)
    End If
    
    AddToDataModel = True
    Exit Function
    
ErrorHandler:
    AddToDataModel = False
End Function

Function CreateRelationships() As Boolean
    ' Create relationships between tables
    On Error GoTo ErrorHandler
    
    Dim model As Model
    Set model = ActiveWorkbook.Model
    
    ' Create relationship between ProductList and SimilarityMatrix on Product Index
    model.ModelRelationships.Add _
        model.ModelTables("SimilarityMatrix").ModelColumns("Product Index"), _
        model.ModelTables("ProductList").ModelColumns("Product Index")
    
    CreateRelationships = True
    Exit Function
    
ErrorHandler:
    CreateRelationships = False
End Function

Sub AddCalculatedColumns()
    ' Add calculated columns to the Similarity Matrix table
    On Error GoTo ErrorHandler
    
    Dim model As Model
    Set model = ActiveWorkbook.Model
    
    ' Add Similarity % column
    With model.ModelTables("SimilarityMatrix").AddColumns("Similarity %")
        .Formula = "[Similarity Value] * 100"
        .DataType = xlDataTypeDouble
        .FormatInformation = "0.00"
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding calculated columns: " & Err.Description, vbExclamation
End Sub

Sub AddMeasures()
    ' Add measures to the model
    On Error GoTo ErrorHandler
    
    Dim model As Model
    Set model = ActiveWorkbook.Model
    
    ' Add measure to SimilarityMatrix table
    With model.ModelTables("SimilarityMatrix").AddMeasure("Count Above Threshold")
        .Formula = "=COUNTROWS(FILTER(SimilarityMatrix, [Similarity %] > SELECTEDVALUE(ThresholdSummary[Threshold])))"
        .FormatInformation = "#,##0"
    End With
    
    With model.ModelTables("SimilarityMatrix").AddMeasure("Average Similarity")
        .Formula = "=AVERAGE(SimilarityMatrix[Similarity %])"
        .FormatInformation = "0.00%"
    End With
    
    With model.ModelTables("SimilarityMatrix").AddMeasure("Max Similarity")
        .Formula = "=MAX(SimilarityMatrix[Similarity %])"
        .FormatInformation = "0.00%"
    End With
    
    ' Add measures to ThresholdSummary table
    With model.ModelTables("ThresholdSummary").AddMeasure("Products in Groups %")
        .Formula = "=DIVIDE(SUM(ThresholdSummary[Products in Groups]), CALCULATE(SUM(ThresholdSummary[Products in Groups]), ALL(ThresholdSummary)))"
        .FormatInformation = "0.00%"
    End With
    
    With model.ModelTables("ThresholdSummary").AddMeasure("Avg Group Size")
        .Formula = "=AVERAGE(ThresholdSummary[Products in Groups] / ThresholdSummary[Groups Found])"
        .FormatInformation = "#,##0.0"
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding measures: " & Err.Description, vbExclamation
End Sub

Sub CreateAnalysisSheet()
    ' Create a sheet for custom threshold analysis
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim analysisSheet As Worksheet
    
    ' Check if Analysis sheet already exists
    On Error Resume Next
    Set analysisSheet = ThisWorkbook.Sheets("Analysis")
    On Error GoTo ErrorHandler
    
    If analysisSheet Is Nothing Then
        Set analysisSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        analysisSheet.Name = "Analysis"
    Else
        analysisSheet.Cells.Clear
    End If
    
    ' Set up the analysis sheet
    With analysisSheet
        ' Headers
        .Range("A1").Value = "Threshold Analysis"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Current Threshold:"
        .Range("B3").Value = 75
        .Range("B3").Font.Bold = True
        
        ' Instructions
        .Range("A5").Value = "Instructions:"
        .Range("A6").Value = "1. Change the threshold value in cell B3"
        .Range("A7").Value = "2. Click 'Update Analysis' button"
        .Range("A8").Value = "3. Results will show below"
        
        ' Create button for updating analysis
        .Buttons.Add(100, 20, 120, 30).OnAction = "UpdateAnalysis"
        .Buttons(.Buttons.Count).Text = "Update Analysis"
        
        ' Results area
        .Range("A10").Value = "Analysis Results"
        .Range("A10").Font.Bold = True
        .Range("A10").Font.Size = 14
        
        .Range("A11").Value = "Products Above Threshold:"
        .Range("B11").Formula = "=COUNTIF(SimilarityMatrix!C:Z,">" & B3/100 & ")"
        .Range("B11").Font.Bold = True
        
        .Range("A12").Value = "Potential Groups:"
        .Range("B12").Formula = "=ROUND(B11/2,0)"
        .Range("B12").Font.Bold = True
        
        ' Column widths
        .Columns("A:A").ColumnWidth = 30
        .Columns("B:B").ColumnWidth = 15
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating analysis sheet: " & Err.Description, vbExclamation
End Sub

Sub UpdateAnalysis()
    ' Update analysis based on current threshold
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Analysis")
    
    Dim threshold As Double
    threshold = ws.Range("B3").Value
    
    ' Update formulas
    ws.Range("B11").Formula = "=COUNTIF(SimilarityMatrix!C:Z,">" & threshold/100 & ")"
    ws.Range("B12").Formula = "=ROUND(B11/2,0)"
    
    ' Highlight cells above threshold in similarity matrix
    HighlightSimilarityMatrix threshold
    
    MsgBox "Analysis updated for threshold " & threshold & "%", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating analysis: " & Err.Description, vbExclamation
End Sub

Sub HighlightSimilarityMatrix(threshold As Double)
    ' Apply conditional formatting to similarity matrix
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Similarity Matrix")
    
    ' Clear existing conditional formatting
    ws.Cells.FormatConditions.Delete
    
    ' Find the data range
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Apply conditional formatting for values above threshold
    With ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, lastCol))
        .FormatConditions.AddColorScale ColorScaleType:=3
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)
        .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .FormatConditions(1).ColorScaleCriteria(2).Value = threshold / 100
        .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0)
        .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0)
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error highlighting similarity matrix: " & Err.Description, vbExclamation
End Sub

Sub CreateDashboard()
    ' Create a dashboard with PivotTable and charts
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim dashboardSheet As Worksheet
    
    ' Check if Dashboard sheet already exists
    On Error Resume Next
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo ErrorHandler
    
    If dashboardSheet Is Nothing Then
        Set dashboardSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        dashboardSheet.Name = "Dashboard"
    Else
        dashboardSheet.Cells.Clear
    End If
    
    ' Create PivotTable
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    Set pc = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlExternal, _
        SourceData:=ActiveWorkbook.ModelConnection)
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=dashboardSheet.Range("A5"), _
        TableName:="ThresholdPivot")
    
    ' Configure PivotTable
    With pt
        ' Add fields
        .PivotFields("Threshold").Orientation = xlRowField
        .PivotFields("Groups Found").Orientation = xlDataField
        .PivotFields("Products in Groups").Orientation = xlDataField
        .PivotFields("Largest Group Size").Orientation = xlDataField
        
        ' Add slicer
        .PivotFields("Threshold").Parent.Slicers.Add(dashboardSheet, "Threshold", "Threshold", "Threshold Slicer", 100, 100, 150, 200)
    End With
    
    ' Create chart
    Dim ch As ChartObject
    Set ch = dashboardSheet.ChartObjects.Add(Left:=300, Top:=50, Width:=400, Height:=300)
    
    With ch.Chart
        .SetSourceData Source:=pt.TableRange2
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Groups vs Threshold"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Threshold (%)"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Count"
    End With
    
    ' Add title
    dashboardSheet.Range("A1").Value = "Threshold Explorer Dashboard"
    dashboardSheet.Range("A1").Font.Size = 20
    dashboardSheet.Range("A1").Font.Bold = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating dashboard: " & Err.Description, vbExclamation
End Sub

Sub AddInstructions()
    ' Add an instructions sheet
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim instructionSheet As Worksheet
    
    ' Check if Instructions sheet already exists
    On Error Resume Next
    Set instructionSheet = ThisWorkbook.Sheets("Instructions")
    On Error GoTo ErrorHandler
    
    If instructionSheet Is Nothing Then
        Set instructionSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        instructionSheet.Name = "Instructions"
    Else
        instructionSheet.Cells.Clear
    End If
    
    ' Add instructions
    With instructionSheet
        .Range("A1").Value = "Quick Reference Guide"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "How to Use This Workbook:"
        .Range("A3").Font.Bold = True
        .Range("A3").Font.Size = 14
        
        .Range("A5").Value = "1. Analysis Sheet:"
        .Range("B5").Value = "Test custom thresholds and see immediate results"
        
        .Range("A6").Value = "2. Dashboard Sheet:"
        .Range("B6").Value = "Interactive charts with threshold slicer"
        
        .Range("A7").Value = "3. Power Pivot:"
        .Range("B7").Value = "Advanced analysis with DAX measures"
        
        .Range("A9").Value = "Tips:"
        .Range("A9").Font.Bold = True
        
        .Range("A10").Value = "• Use the slicer on Dashboard to explore thresholds"
        .Range("A11").Value = "• Green cells in Similarity Matrix are above threshold"
        .Range("A12").Value = "• Create new PivotTables from the Data Model"
        .Range("A13").Value = "• Modify measures in Power Pivot for custom calculations"
        
        .Range("A15").Value = "Keyboard Shortcuts:"
        .Range("A15").Font.Bold = True
        
        .Range("A16").Value = "• Alt+F1: Create chart from data"
        .Range("A17").Value = "• Alt+N+V: Create PivotTable"
        .Range("A18").Value = "• Alt+P+C: Manage Power Pivot"
        
        ' Column widths
        .Columns("A:A").ColumnWidth = 25
        .Columns("B:B").ColumnWidth = 50
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding instructions: " & Err.Description, vbExclamation
End Sub

'================================================================================================
' Additional utility macros
'================================================================================================

Sub ExportCurrentThreshold()
    ' Export the current threshold analysis to a new workbook
    On Error GoTo ErrorHandler
    
    Dim threshold As Double
    threshold = ThisWorkbook.Sheets("Analysis").Range("B3").Value
    
    Dim newWb As Workbook
    Set newWb = Workbooks.Add
    
    ' Copy relevant sheets
    ThisWorkbook.Sheets("Threshold Summary").Copy Before:=newWb.Sheets(1)
    ThisWorkbook.Sheets("Threshold Groups").Copy Before:=newWb.Sheets(1)
    
    ' Filter for current threshold
    newWb.Sheets("Threshold Groups").AutoFilter Field:=1, Criteria1:=threshold
    
    MsgBox "Exported threshold " & threshold & "% analysis to new workbook", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting: " & Err.Description, vbExclamation
End Sub

Sub ResetWorkbook()
    ' Reset the workbook to original state
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    
    ' Delete analysis sheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Analysis" Or ws.Name = "Dashboard" Or ws.Name = "Instructions" Then
            ws.Delete
        End If
    Next ws
    
    ' Clear Power Pivot model (optional - comment out if you want to keep it)
    ' ActiveWorkbook.Model.Initialize
    
    Application.DisplayAlerts = True
    
    MsgBox "Workbook reset to original state", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "Error resetting: " & Err.Description, vbExclamation
End Sub
