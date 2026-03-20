# Excel VBA Macro Setup Guide

This guide shows how to use the VBA macro to automate the Excel Threshold Explorer setup.

## Easiest Option: Use the Ready-to-Use Excel File

We've created a ready-to-use Excel file that contains the macro code in a worksheet:

1. Open `Threshold_Explorer_Ready_to_Use.xlsx` (included in the repository)
2. Follow the instructions on the first sheet
3. Copy the VBA code from column B
4. Paste into the VBA editor

This is the simplest method - no need to open separate files!

## Alternative: Manual Import

If you prefer to import the macro file directly:

## What the Macro Does

The `SetupThresholdExplorer` macro automatically:
1. Adds all data tables to Power Pivot
2. Creates relationships between tables
3. Adds calculated columns (Similarity %)
4. Creates DAX measures for analysis
5. Sets up an Analysis sheet for custom threshold testing
6. Creates a Dashboard with charts and slicers
7. Adds an Instructions sheet for quick reference

## How to Install and Run the Macro

### Step 1: Enable the Developer Tab in Excel
1. Click **File** → **Options** → **Customize Ribbon**
2. Check the **Developer** box in the right panel
3. Click **OK**

### Step 2: Enable Macros (if needed)
1. Click **File** → **Options** → **Trust Center**
2. Click **Trust Center Settings**
3. Select **Macro Settings**
4. Choose **Enable all macros** (for testing) or **Disable all macros with notification**
5. Click **OK** twice

### Step 3: Open the VBA Editor
1. Open your `enhanced_threshold_analysis.xlsx` file
2. Press **Alt+F11** or click **Developer** → **Visual Basic**
3. The VBA editor will open

### Step 4: Import the Macro
1. In the VBA editor, go to **File** → **Import File**
2. Navigate to and select `Excel_Threshold_Explorer_Macro.bas`
3. The macro will be imported into a new module

### Step 5: Run the Macro
1. With the VBA editor still open:
   - Click anywhere inside the `SetupThresholdExplorer` subroutine
   - Press **F5** to run
   - OR click **Run** → **Run Sub/UserForm**
   - OR press the green play button

2. Alternatively, from Excel:
   - Press **Alt+F8** to open the Macro dialog
   - Select `SetupThresholdExplorer`
   - Click **Run**

### Step 6: Save the File with Macros
1. After the macro runs, save your file:
   - Click **File** → **Save As**
   - Choose **Excel Macro-Enabled Workbook (*.xlsm)**
   - Save the file

## What You'll Get After Running the Macro

### New Sheets Created:
1. **Analysis** - Test custom thresholds with instant results
2. **Dashboard** - Interactive charts with threshold slicer
3. **Instructions** - Quick reference guide

### Power Pivot Setup:
- All tables added to the data model
- Relationships created automatically
- Calculated columns and measures added

### Visual Enhancements:
- Conditional formatting on the similarity matrix
- Color-coded cells (red=low, yellow=medium, green=high similarity)

## Using the Created Sheets

### Analysis Sheet:
1. Change the threshold value in cell B3
2. Click "Update Analysis" button
3. See results update instantly
4. Similarity matrix highlights cells above threshold

### Dashboard Sheet:
1. Use the threshold slicer to filter data
2. Charts update automatically
3. Hover over data points for details

### Power Pivot (Advanced):
1. Press **Alt+P+C** to open Power Pivot
2. View the data model diagram
3. Modify or add measures as needed

## Additional Macros Included

- **UpdateAnalysis** - Updates the analysis sheet with new threshold
- **ExportCurrentThreshold** - Exports analysis for current threshold
- **ResetWorkbook** - Removes all added sheets and resets

## Troubleshooting

### "Macro not found" error:
- Ensure you've imported the `.bas` file correctly
- Check that macros are enabled in Excel Trust Center

### "Power Pivot not available" error:
- You need Excel 2016+ or Excel 2013 with Power Pivot add-in
- Enable Power Pivot: **File** → **Options** → **Add-Ins** → Manage: COM Add-ins → Check "Power Pivot for Excel"

### "Runtime error" during setup:
- Check that all required sheets exist in your workbook
- Ensure your data has the expected column headers
- Try running with a smaller dataset first

## Security Note

Only enable macros from trusted sources. This macro is provided as-is and you should review the code before running it in a production environment.

## Customization Tips

You can modify the macro to:
- Change default threshold value
- Add more calculated columns
- Create additional charts
- Customize formatting styles

Edit the macro in the VBA editor (Alt+F11) and make your changes, then run again.
