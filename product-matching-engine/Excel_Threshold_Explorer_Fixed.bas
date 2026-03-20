Option Explicit

'================================================================================================
' Product Matching Engine - Excel Threshold Explorer - Fixed Version
'================================================================================================
' This version checks for Power Pivot availability and provides clear error messages
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
    
    ' Step 1: Check if Power Pivot is available
    If Not IsPowerPivotAvailable() Then
        MsgBox "Power Pivot is not available or not enabled." & vbCrLf & vbCrLf & _
               "To enable Power Pivot:" & vbCrLf & _
               "1. Go to File > Options > Add-ins" & vbCrLf & _
               "2. Select 'COM Add-ins' and click Go" & vbCrLf & _
               "3. Check 'Microsoft Power Pivot for Excel'" & vbCrLf & _
               "4. Click OK and restart Excel" & vbCrLf & vbCrLf & _
               "Continuing with basic setup (no Power Pivot features)...", _
               vbInformation, "Power Pivot Not Available"
        
        ' Create basic analysis without Power Pivot
        CreateBasicAnalysisSheet
        GoTo CleanupSuccess
    End If

    ' Step 2: Add tables to Power Pivot
    If Not AddTablesToModel() Then
        GoTo CleanupFail
    End If
    
    ' Step 3: Create relationships
    If Not CreateRelationships() Then
        GoTo CleanupFail
    End If
    
    ' Step 4: Add calculated columns and measures
    If Not AddCalculatedColumnsAndMeasures() Then
        GoTo CleanupFail
    End If
    
    ' Step 5: Create analysis sheet
    CreateAnalysisSheet
    
    ' Step 6: Create dashboard
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
    MsgBox "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Line: " & Erl, vbCritical, "Error"
End Sub

Sub RunStep0_ValidateSheets()
    If ValidateRequiredSheets() Then
        MsgBox "Step 0 complete: required sheets are present.", vbInformation
    End If
End Sub

Sub RunStep1_CheckPowerPivot()
    If IsPowerPivotAvailable() Then
        MsgBox "Step 1 complete: Power Pivot is available.", vbInformation
    Else
        MsgBox "Step 1 failed: Power Pivot is not available/enabled.", vbExclamation
    End If
End Sub

Sub RunStep2_AddTablesToModel()
    If AddTablesToModel() Then
        MsgBox "Step 2 complete: tables added to Data Model.", vbInformation
    End If
End Sub

Sub RunStep3_CreateRelationships()
    If CreateRelationships() Then
        MsgBox "Step 3 complete: relationships created.", vbInformation
    End If
End Sub

Sub RunStep4_AddCalculatedColumnsAndMeasures()
    If AddCalculatedColumnsAndMeasures() Then
        MsgBox "Step 4 complete: calculated columns and measures added.", vbInformation
    End If
End Sub

Sub RunStep5_CreateAnalysisSheet()
    CreateAnalysisSheet
    MsgBox "Step 5 complete: Analysis sheet created.", vbInformation
End Sub

Sub RunStep6_CreateDashboard()
    CreateDashboard
    MsgBox "Step 6 complete: Dashboard sheet created.", vbInformation
End Sub

Sub RunManualFallbackGuide()
    Dim msg As String
    msg = "Manual fallback guide:" & vbCrLf & vbCrLf & _
          "1) Add Product List / Similarity Matrix tables to Data Model manually:" & vbCrLf & _
          "   Select each table > Power Pivot tab > Add to Data Model" & vbCrLf & vbCrLf & _
          "2) Create relationship manually in Power Pivot Manage:" & vbCrLf & _
          "   Similarity Matrix[Product Index] -> Product List[Product Index]" & vbCrLf & vbCrLf & _
          "3) Add DAX manually if needed:" & vbCrLf & _
          "   Similarity % = [Similarity Value] * 100" & vbCrLf & _
          "   Measure: Average Similarity = AVERAGE('Similarity Matrix'[Similarity %])"
    MsgBox msg, vbInformation, "Threshold Explorer - Manual Fallback"
End Sub

Function IsPowerPivotAvailable() As Boolean
    ' Check if Power Pivot is available
    On Error Resume Next
    
    Dim model As Object
    Set model = ThisWorkbook.model
    
    If Err.Number <> 0 Then
        IsPowerPivotAvailable = False
    Else
        IsPowerPivotAvailable = Not (model Is Nothing)
    End If
    
    On Error GoTo 0
End Function

Function AddTablesToModel() As Boolean
    ' Add all data tables to Power Pivot model
    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error GoTo ErrorHandler
    
    ' Add Product List
    Set ws = GetSheet("Product List")
    If ws Is Nothing Then
        MsgBox "Product List sheet not found", vbExclamation
        AddTablesToModel = False
        Exit Function
    End If
    Set tbl = EnsureWorksheetTable(ws, "ProductList")
    If tbl Is Nothing Then
        AddTablesToModel = False
        Exit Function
    End If

    If Not AddTableToDataModel(tbl, "ProductList") Then
        AddTablesToModel = False
        Exit Function
    End If
    
    ' Add Similarity Matrix
    Set ws = GetSheet("Similarity Matrix")
    If ws Is Nothing Then
        MsgBox "Similarity Matrix sheet not found", vbExclamation
        AddTablesToModel = False
        Exit Function
    End If
    Set tbl = EnsureWorksheetTable(ws, "SimilarityMatrix")
    If tbl Is Nothing Then
        AddTablesToModel = False
        Exit Function
    End If

    If Not AddTableToDataModel(tbl, "SimilarityMatrix") Then
        AddTablesToModel = False
        Exit Function
    End If
    
    ' Add Threshold Summary if exists
    Set ws = GetSheet("Threshold Summary")
    If Not ws Is Nothing Then
        Set tbl = EnsureWorksheetTable(ws, "ThresholdSummary")
        If Not tbl Is Nothing Then
            If Not AddTableToDataModel(tbl, "ThresholdSummary") Then
                AddTablesToModel = False
                Exit Function
            End If
        End If
    End If
    
    AddTablesToModel = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to add tables to Power Pivot." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Number: " & Err.Number, _
           vbExclamation, "Power Pivot Error"
    AddTablesToModel = False
End Function

Function EnsureWorksheetTable(ws As Worksheet, tableName As String) As ListObject
    Dim tbl As ListObject
    Dim dataRange As Range

    On Error GoTo ErrorHandler

    Set dataRange = ws.UsedRange
    If dataRange Is Nothing Then
        MsgBox ws.Name & " sheet has no used range.", vbExclamation
        Set EnsureWorksheetTable = Nothing
        Exit Function
    End If

    If dataRange.Rows.Count < 2 Or dataRange.Columns.Count < 1 Then
        MsgBox ws.Name & " sheet does not contain a header row and data rows.", vbExclamation
        Set EnsureWorksheetTable = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo ErrorHandler

    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        tbl.Name = tableName
    Else
        tbl.Resize dataRange
    End If

    tbl.TableStyle = "TableStyleMedium2"
    Set EnsureWorksheetTable = tbl
    Exit Function

ErrorHandler:
    MsgBox "Failed to prepare Excel Table '" & tableName & "' on sheet '" & ws.Name & "': " & Err.Description & " (" & Err.Number & ")", vbExclamation
    Set EnsureWorksheetTable = Nothing
End Function

Function AddTableToDataModel(tbl As ListObject, baseName As String) As Boolean
    Dim ok As Boolean
    Dim errMsg As String

    ok = TryAddViaExecuteMso(tbl, errMsg)
    If ok Then
        AddTableToDataModel = True
        Exit Function
    End If

    ok = TryAddViaConnectionsAdd2(tbl, errMsg)
    If ok Then
        AddTableToDataModel = True
        Exit Function
    End If

    MsgBox "Failed to add table '" & tbl.Name & "' to Data Model." & vbCrLf & vbCrLf & errMsg, vbExclamation
    AddTableToDataModel = False
End Function

Function TryAddViaExecuteMso(tbl As ListObject, ByRef errMsg As String) As Boolean
    On Error GoTo ErrorHandler

    tbl.Parent.Activate
    tbl.Range.Cells(1, 1).Select
    Application.CommandBars.ExecuteMso "PowerPivotAddToDataModel"
    DoEvents

    TryAddViaExecuteMso = True
    Exit Function

ErrorHandler:
    errMsg = "ExecuteMso failed: " & Err.Description & " (" & Err.Number & ")"
    TryAddViaExecuteMso = False
End Function

Function TryAddViaConnectionsAdd2(tbl As ListObject, ByRef errMsg As String) As Boolean
    Dim connName As String
    Dim connString As String
    Dim cmdText As String
    Dim conn As WorkbookConnection

    On Error GoTo ErrorHandler

    connName = "WorksheetConnection_" & ThisWorkbook.Name & "!" & tbl.Name

    ' If model connection already exists, treat as success.
    For Each conn In ThisWorkbook.Connections
        If conn.Name = connName Then
            TryAddViaConnectionsAdd2 = True
            Exit Function
        End If
    Next conn

    connString = "WORKSHEET;" & ThisWorkbook.Path & "\" & ThisWorkbook.Name
    cmdText = ThisWorkbook.Name & "!" & tbl.Name

    ThisWorkbook.Connections.Add2 _
        Name:=connName, _
        Description:="", _
        ConnectionString:=connString, _
        CommandText:=cmdText, _
        lCmdtype:=7, _
        CreateModelConnection:=True, _
        ImportRelationships:=False

    TryAddViaConnectionsAdd2 = True
    Exit Function

ErrorHandler:
    errMsg = errMsg & vbCrLf & "Add2 failed: " & Err.Description & " (" & Err.Number & ")"
    TryAddViaConnectionsAdd2 = False
End Function

Function CreateRelationships() As Boolean
    ' Create relationship between tables
    Dim model As Object
    Dim simTable As Object, prodTable As Object
    Dim simCol As Object, prodCol As Object
    On Error GoTo ErrorHandler
    Set model = ThisWorkbook.model
    
    ' Check if relationship already exists
    On Error Resume Next
    Dim rel As Object
    For Each rel In model.ModelRelationships
        If rel.ForeignKeyColumn.Name = "Product Index" And _
           rel.PrimaryKeyColumn.Name = "Product Index" Then
            CreateRelationships = True
            Exit Function
        End If
    Next rel
    On Error GoTo ErrorHandler
    
    ' Resolve tables and key columns with name normalization
    Set simTable = ResolveModelTable(model, "SimilarityMatrix", "Similarity Matrix")
    Set prodTable = ResolveModelTable(model, "ProductList", "Product List")

    If simTable Is Nothing Then
        MsgBox "Similarity table not found in Data Model." & vbCrLf & vbCrLf & _
               "Model tables: " & GetModelTableNames(model), vbExclamation, "Table Not Found"
        CreateRelationships = False
        Exit Function
    End If

    If prodTable Is Nothing Then
        MsgBox "Product table not found in Data Model." & vbCrLf & vbCrLf & _
               "Model tables: " & GetModelTableNames(model), vbExclamation, "Table Not Found"
        CreateRelationships = False
        Exit Function
    End If

    Set simCol = ResolveModelColumn(simTable, "Product Index")
    Set prodCol = ResolveModelColumn(prodTable, "Product Index")
    
    If simCol Is Nothing Then
        MsgBox "Column 'Product Index' not found in Similarity Matrix table." & vbCrLf & vbCrLf & _
               "Available columns:" & vbCrLf & _
               GetColumnNames("SimilarityMatrix"), _
               vbExclamation, "Column Not Found"
        CreateRelationships = False
        Exit Function
    End If
    
    If prodCol Is Nothing Then
        MsgBox "Column 'Product Index' not found in Product List table." & vbCrLf & vbCrLf & _
               "Available columns:" & vbCrLf & _
               GetColumnNames("ProductList"), _
               vbExclamation, "Column Not Found"
        CreateRelationships = False
        Exit Function
    End If
    
    ' Create the relationship
    model.ModelRelationships.Add simCol, prodCol

    CreateRelationships = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to create relationships." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Number: " & Err.Number, _
           vbExclamation, "Relationship Error"
    CreateRelationships = False
End Function

Function GetColumnNames(tableName As String) As String
    ' Get list of column names from a Power Pivot model table
    On Error Resume Next

    Dim model As Object
    Set model = ThisWorkbook.model
    
    If model Is Nothing Then
        GetColumnNames = "Power Pivot model not available"
        Exit Function
    End If
    
    Dim tbl As Object
    Set tbl = ResolveModelTable(model, tableName, Replace(tableName, " ", ""))
    If tbl Is Nothing Then
        If LCase$(tableName) = "similaritymatrix" Then Set tbl = ResolveModelTable(model, "Similarity Matrix", "SimilarityMatrix")
        If LCase$(tableName) = "productlist" Then Set tbl = ResolveModelTable(model, "Product List", "ProductList")
    End If
    
    If tbl Is Nothing Then
        ' Try getting from worksheet instead
        Dim ws As Worksheet
        Set ws = GetSheet(tableName)
        If ws Is Nothing Then
            If LCase$(tableName) = "similaritymatrix" Then Set ws = GetSheet("Similarity Matrix")
            If LCase$(tableName) = "productlist" Then Set ws = GetSheet("Product List")
        End If
        
        If ws Is Nothing Then
            GetColumnNames = "Table/Sheet not found"
            Exit Function
        End If
        
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        Dim i As Long
        Dim colNames As String
        For i = 1 To lastCol
            If i > 1 Then colNames = colNames & ", "
            colNames = colNames & ws.Cells(1, i).Value
        Next i
        
        GetColumnNames = colNames
        Exit Function
    End If
    
    ' Get columns from Power Pivot table
    Dim col As Object
    Dim colList As String
    For Each col In tbl.ModelColumns
        If colList <> "" Then colList = colList & ", "
        colList = colList & col.Name
    Next col
    
    GetColumnNames = colList
    On Error GoTo 0
End Function

Function ResolveModelTable(model As Object, preferredName As String, alternateName As String) As Object
    Dim t As Object
    Dim nPref As String, nAlt As String, nName As String, nSource As String

    nPref = NormalizeName(preferredName)
    nAlt = NormalizeName(alternateName)

    On Error Resume Next
    Set ResolveModelTable = model.ModelTables(preferredName)
    If Not ResolveModelTable Is Nothing Then Exit Function
    Set ResolveModelTable = model.ModelTables(alternateName)
    If Not ResolveModelTable Is Nothing Then Exit Function
    On Error GoTo 0

    For Each t In model.ModelTables
        nName = NormalizeName(t.Name)
        nSource = NormalizeName(t.SourceName)
        If nName = nPref Or nName = nAlt Or nSource = nPref Or nSource = nAlt Then
            Set ResolveModelTable = t
            Exit Function
        End If
    Next t

    Set ResolveModelTable = Nothing
End Function

Function ResolveModelColumn(modelTable As Object, desiredName As String) As Object
    Dim c As Object
    Dim target As String
    target = NormalizeName(desiredName)

    On Error Resume Next
    Set ResolveModelColumn = modelTable.ModelColumns(desiredName)
    If Not ResolveModelColumn Is Nothing Then Exit Function
    On Error GoTo 0

    For Each c In modelTable.ModelColumns
        If NormalizeName(c.Name) = target Then
            Set ResolveModelColumn = c
            Exit Function
        End If
    Next c

    Set ResolveModelColumn = Nothing
End Function

Function GetModelTableNames(model As Object) As String
    Dim t As Object
    Dim names As String
    For Each t In model.ModelTables
        If names <> "" Then names = names & ", "
        names = names & t.Name
    Next t
    GetModelTableNames = names
End Function

Function NormalizeName(value As String) As String
    NormalizeName = LCase$(Replace(Replace(Replace(Trim$(value), " ", ""), "_", ""), "-", ""))
End Function

Function AddCalculatedColumnsAndMeasures() As Boolean
    ' Add calculated columns and DAX measures
    Dim model As Object
    Dim simTable As Object
    Dim simTableName As String
    On Error GoTo ErrorHandler
    Set model = ThisWorkbook.model

    Set simTable = ResolveModelTable(model, "SimilarityMatrix", "Similarity Matrix")
    If simTable Is Nothing Then
        MsgBox "Similarity table not found in Data Model." & vbCrLf & _
               "Model tables: " & GetModelTableNames(model), vbExclamation, "DAX Error"
        AddCalculatedColumnsAndMeasures = False
        Exit Function
    End If
    simTableName = simTable.Name
    
    ' Check if column already exists
    On Error Resume Next
    Dim col As Object
    Set col = ResolveModelColumn(simTable, "Similarity %")
    On Error GoTo ErrorHandler
    
    ' Add Similarity % column if it doesn't exist
    If col Is Nothing Then
        With model.ModelTables(simTableName).ModelTableColumns.Add("Similarity %")
            .Formula = "[Similarity Value] * 100"
            .DataType = 2 ' xlDataTypeDouble
        End With
    End If
    
    ' Add measures if they don't exist
    On Error Resume Next
    Dim meas As Object
    Set meas = model.ModelTables(simTableName).ModelMeasures("Average Similarity")
    On Error GoTo ErrorHandler
    
    If meas Is Nothing Then
        With model.ModelTables(simTableName).ModelMeasures.Add("Average Similarity", "=AVERAGE('" & simTableName & "'[Similarity %])")
            .FormatInformation = "0.00%"
        End With
    End If
    
    AddCalculatedColumnsAndMeasures = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to add calculated columns/measures." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbExclamation, "DAX Error"
    AddCalculatedColumnsAndMeasures = False
End Function

Sub CreateBasicAnalysisSheet()
    ' Create analysis sheet without Power Pivot
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Analysis")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Analysis"
    
    With ws
        .Range("A1").Value = "Threshold Analysis (Basic Mode)"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Enter Threshold (%):"
        .Range("B3").Value = 75
        .Range("B3").Font.Bold = True
        
        .Range("A5").Value = "Products Above Threshold:"
        If GetSheet("Similarity Matrix") Is Nothing Then
            .Range("B5").Value = "Error: Similarity Matrix sheet not found"
        Else
            .Range("B5").Formula = "=COUNTIF('Similarity Matrix'!C:Z,"">"" & B3/100)"
        End If
        
        .Range("A6").Value = "Potential Groups:"
        .Range("B6").Formula = "=ROUND(B5/2,0)"
        
        .Range("A8").Value = "Note: Power Pivot features not available"
        .Range("A8").Font.Italic = True
        .Range("A8").Font.Color = RGB(255, 0, 0)
        
        .Columns("A:A").ColumnWidth = 30
        .Columns("B:B").ColumnWidth = 15
    End With
End Sub

Sub CreateAnalysisSheet()
    ' Create analysis sheet with Power Pivot
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
            .Range("B5").Formula = "=COUNTIF('Similarity Matrix'!C:Z,"">"" & B3/100)"
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
    ' Create dashboard with PivotTable
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
        
        ' Add chart
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
    
    ws.Range("B5").Formula = "=COUNTIF('Similarity Matrix'!C:Z,"">"" & " & threshold / 100 & ")"
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
