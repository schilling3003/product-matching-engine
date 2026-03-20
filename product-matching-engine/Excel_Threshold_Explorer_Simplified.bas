Option Explicit

'================================================================================================
' Product Matching Engine - Excel Threshold Explorer - Simplified Macro
'================================================================================================
' This is a simplified version with only the essential macros needed.
' Run SetupThresholdExplorer to set up everything automatically.
'================================================================================================

Sub SetupThresholdExplorer()
    ' Main macro - sets up the complete threshold explorer
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Step 0: Validate required sheets
    If Not ValidateRequiredSheets() Then
        GoTo CleanupFail
    End If

    ' Step 1: Add tables to Power Pivot
    If Not AddTablesToModel() Then
        GoTo CleanupFail
    End If
    
    ' Step 2: Create relationships
    If Not CreateRelationships() Then
        GoTo CleanupFail
    End If
    
    ' Step 3: Add calculated columns and measures
    If Not AddCalculatedColumnsAndMeasures() Then
        GoTo CleanupFail
    End If
    
    ' Step 4: Create analysis sheet
    CreateAnalysisSheet
    
    ' Step 5: Create dashboard
    CreateDashboard
    
CleanupSuccess:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Threshold Explorer setup complete!" & vbCrLf & vbCrLf & _
           "Created:" & vbCrLf & _
           "- Analysis sheet (for testing thresholds)" & vbCrLf & _
           "- Dashboard (with charts)", _
           vbInformation, "Complete"
    Exit Sub

CleanupFail:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Function AddTablesToModel() As Boolean
    ' Add all data tables to Power Pivot model
    Dim ws As Worksheet
    Dim model As Model

    On Error GoTo ErrorHandler
    Set model = ThisWorkbook.Model
    
    ' Add Product List
    Set ws = GetSheet("Product List")
    If ws Is Nothing Then GoTo ErrorHandler
    ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = "ProductList"
    model.ModelTables.Add "ProductList", ws.ListObjects(1)
    
    ' Add Similarity Matrix
    Set ws = GetSheet("Similarity Matrix")
    If ws Is Nothing Then GoTo ErrorHandler
    ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = "SimilarityMatrix"
    model.ModelTables.Add "SimilarityMatrix", ws.ListObjects(1)
    
    ' Add Threshold Summary if exists
    On Error Resume Next
    Set ws = GetSheet("Threshold Summary")
    If Not ws Is Nothing Then
        ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = "ThresholdSummary"
        model.ModelTables.Add "ThresholdSummary", ws.ListObjects(1)
    End If
    On Error GoTo 0
    AddTablesToModel = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to add tables to Power Pivot data model. Ensure Power Pivot is enabled and sheets exist.", vbExclamation
    AddTablesToModel = False
End Function

Function CreateRelationships() As Boolean
    ' Create relationship between tables
    Dim model As Model
    On Error GoTo ErrorHandler
    Set model = ThisWorkbook.Model
    
    model.ModelRelationships.Add _
        model.ModelTables("SimilarityMatrix").ModelColumns("Product Index"), _
        model.ModelTables("ProductList").ModelColumns("Product Index")

    CreateRelationships = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to create Power Pivot relationships. Check that 'Product Index' exists in both tables.", vbExclamation
    CreateRelationships = False
End Function

Function AddCalculatedColumnsAndMeasures() As Boolean
    ' Add calculated columns and DAX measures
    Dim model As Model
    On Error GoTo ErrorHandler
    Set model = ThisWorkbook.Model
    
    ' Add Similarity % column
    With model.ModelTables("SimilarityMatrix").AddColumns("Similarity %")
        .Formula = "[Similarity Value] * 100"
        .DataType = xlDataTypeDouble
        .FormatInformation = "0.00"
    End With
    
    ' Add key measures
    With model.ModelTables("SimilarityMatrix").AddMeasure("Count Above Threshold")
        .Formula = "=COUNTROWS(FILTER(SimilarityMatrix, [Similarity %] > SELECTEDVALUE(ThresholdSummary[Threshold])))"
        .FormatInformation = "#,##0"
    End With
    
    With model.ModelTables("SimilarityMatrix").AddMeasure("Average Similarity")
        .Formula = "=AVERAGE(SimilarityMatrix[Similarity %])"
        .FormatInformation = "0.00%"
    End With
    AddCalculatedColumnsAndMeasures = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to add calculated columns/measures. Ensure Threshold Summary exists for measures.", vbExclamation
    AddCalculatedColumnsAndMeasures = False
End Function

Sub CreateAnalysisSheet()
    ' Create simple analysis sheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Analysis")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Analysis"
    
    With ws
        .Range("A1").Value = "Threshold Analysis"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Enter Threshold (%):"
        .Range("B3").Value = 75
        .Range("B3").Font.Bold = True
        
        .Range("A5").Value = "Products Above Threshold:"
        If GetSheet("Similarity Matrix") Is Nothing Then
            .Range("B5").Value = "Error: Similarity Matrix sheet not found"
        Else
            .Range("B5").Formula = "=COUNTIF(SimilarityMatrix!C:Z, "">"" & B3/100)"
        End If
        
        .Range("A6").Value = "Potential Groups:"
        .Range("B6").Formula = "=ROUND(B5/2,0)"
        
        ' Add update button
        .Buttons.Add(150, 50, 100, 30).OnAction = "UpdateAnalysis"
        .Buttons(.Buttons.Count).Text = "Update"
        
        .Columns("A:A").ColumnWidth = 25
        .Columns("B:B").ColumnWidth = 15
    End With
End Sub

Sub CreateDashboard()
    ' Create simple dashboard with PivotTable
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Dashboard")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Dashboard"
    
    ' Create PivotTable from Threshold Summary
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim summarySheet As Worksheet

    Set summarySheet = GetSheet("Threshold Summary")
    If summarySheet Is Nothing Then
        ws.Range("A3").Value = "Threshold Summary sheet not found. Dashboard skipped."
        Exit Sub
    End If

    On Error Resume Next
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=summarySheet.UsedRange)
    
    If Not pc Is Nothing Then
        Set pt = pc.CreatePivotTable( _
            TableDestination:=ws.Range("A5"), _
            TableName:="ThresholdPivot")
        
        With pt
            .PivotFields("Threshold").Orientation = xlRowField
            .PivotFields("Groups Found").Orientation = xlDataField
            .PivotFields("Products in Groups").Orientation = xlDataField
        End With
        
        ' Add simple chart
        Dim ch As ChartObject
        Set ch = ws.ChartObjects.Add(Left:=200, Top:=50, Width:=300, Height:=200)
        ch.Chart.SetSourceData Source:=pt.TableRange2
        ch.Chart.ChartType = xlLine
        ch.Chart.HasTitle = True
        ch.Chart.ChartTitle.Text = "Groups vs Threshold"
    End If
    On Error GoTo 0
    
    ws.Range("A1").Value = "Threshold Dashboard"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
End Sub

Sub UpdateAnalysis()
    ' Update analysis when threshold changes
    Dim ws As Worksheet
    Set ws = GetSheet("Analysis")
    If ws Is Nothing Then
        MsgBox "Analysis sheet not found. Run SetupThresholdExplorer first.", vbExclamation
        Exit Sub
    End If

    If GetSheet("Similarity Matrix") Is Nothing Then
        MsgBox "Similarity Matrix sheet not found.", vbExclamation
        Exit Sub
    End If
    
    Dim threshold As Double
    threshold = ws.Range("B3").Value
    
    ws.Range("B5").Formula = "=COUNTIF(SimilarityMatrix!C:Z,"">"" & " & threshold / 100 & ")"
    ws.Range("B6").Formula = "=ROUND(B5/2,0)"
    
    ' Highlight matrix
    HighlightMatrix threshold
    
    MsgBox "Updated for threshold " & threshold & "%", vbInformation
End Sub

Sub HighlightMatrix(threshold As Double)
    ' Apply conditional formatting to similarity matrix
    Dim ws As Worksheet
    Set ws = GetSheet("Similarity Matrix")
    If ws Is Nothing Then
        MsgBox "Similarity Matrix sheet not found.", vbExclamation
        Exit Sub
    End If
    
    ws.Cells.FormatConditions.Delete
    
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    With ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, lastCol))
        .FormatConditions.AddColorScale ColorScaleType:=3
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 200, 200)
        .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .FormatConditions(1).ColorScaleCriteria(2).Value = threshold / 100
        .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 200)
        .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(200, 255, 200)
    End With
End Sub

Function ValidateRequiredSheets() As Boolean
    If GetSheet("Product List") Is Nothing Then
        MsgBox "Missing sheet: Product List", vbExclamation
        ValidateRequiredSheets = False
        Exit Function
    End If
    If GetSheet("Similarity Matrix") Is Nothing Then
        MsgBox "Missing sheet: Similarity Matrix", vbExclamation
        ValidateRequiredSheets = False
        Exit Function
    End If
    ValidateRequiredSheets = True
End Function

Function GetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
End Function
