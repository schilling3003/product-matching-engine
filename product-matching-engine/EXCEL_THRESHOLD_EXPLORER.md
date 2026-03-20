# Excel Threshold Explorer Guide

This guide shows how to recreate the Threshold Explorer functionality in Excel using the enhanced export data.

## What's Included in the Enhanced Excel Export?

The `enhanced_threshold_analysis.xlsx` file contains:

1. **Similarity Matrix**: A matrix showing similarity scores between all product pairs (sampled to 1000 products for performance)
2. **Product List**: Complete list of products with all their attributes
3. **Threshold Summary**: Summary statistics for each threshold level
4. **Threshold Groups**: Detailed group memberships at each threshold
5. **Excel Instructions**: Step-by-step guide within the file

## Method 1: Using Power Pivot (Recommended for Excel 2016+)

### Setup:
1. Open the Excel file
2. Go to **Power Pivot** tab → **Add to Data Model**
3. Select both "Product List" and "Similarity Matrix" tables
4. Click **Manage** to open the Power Pivot window

### Create Relationships:
1. In the Power Pivot window, go to **Diagram View**
2. Drag "Product Index" from Product List to "Product Index" in Similarity Matrix
3. This creates a relationship between the tables

### Add Calculated Columns:
1. In the Similarity Matrix table, add a calculated column:
   ```
   Similarity % = [Similarity Value] * 100
   ```

### Create Measures:
1. Add the following measures to the Similarity Matrix table:
   ```
   Count Above Threshold = COUNTROWS(
       FILTER(SimilarityMatrix, [Similarity %] > SELECTEDVALUE(ThresholdSummary[Threshold]))
   )
   
   Average Similarity = AVERAGE(SimilarityMatrix[Similarity %])
   
   Groups Count = DISTINCTCOUNT(SimilarityMatrix[Group ID])
   ```

### Create PivotTable:
1. Go back to Excel → **Insert** → **PivotTable**
2. Choose "Use this workbook's Data Model"
3. Add fields:
   - Rows: Product Name
   - Values: Count Above Threshold, Average Similarity
4. Add a **Slicer** for Threshold from the ThresholdSummary table

## Method 2: Using Excel Formulas (All Excel Versions)

### Quick Analysis:
1. Create a new sheet named "Analysis"
2. In cell A1, enter your desired threshold (e.g., 75)
3. Copy product names from Product List to column B
4. In column C, use this formula to count matches above threshold:
   ```
   =COUNTIF(INDEX(SimilarityMatrix!C:ZZ, MATCH(B2, SimilarityMatrix!$B:$B, 0), 0), ">"&$A$1/100)
   ```
5. Drag the formula down for all products

### Visualize with Conditional Formatting:
1. Select the similarity matrix data
2. Go to **Home** → **Conditional Formatting** → **Color Scales**
3. Choose a color scale to visualize similarity patterns

## Method 3: Using Pivot Tables

### Create Dynamic Threshold Analysis:
1. Select the Similarity Matrix data
2. **Insert** → **PivotTable**
3. Configure:
   - Rows: Product Name
   - Values: Average of Similarity Value
   - Columns: (leave empty for now)

### Add Threshold Filtering:
1. Right-click the PivotTable → **Value Filters** → **Top 10**
2. Modify to show values greater than your threshold
3. Or use a slicer connected to the Threshold Summary table

## Advanced Techniques

### 1. Create Dynamic Charts:
```excel
# Create a chart showing how groups change with threshold
=LINEST(ThresholdSummary[Groups Found], ThresholdSummary[Threshold])
```

### 2. Build a Dashboard:
1. Create multiple PivotTables showing different aspects
2. Combine with charts for:
   - Groups vs Threshold line chart
   - Products in Groups vs Threshold
   - Average Similarity by Group Size

### 3. Use Power Query for Large Datasets:
If you have more than 1000 products:
1. **Data** → **Get Data** → **From File** → **From Workbook**
2. Select the similarity matrix
3. Use Power Query to:
   - Filter by threshold
   - Group products
   - Create custom calculations

## Tips for Effective Analysis

1. **Start with the Threshold Summary sheet** to understand the overall pattern
2. **Use conditional formatting** to quickly identify high-similarity regions
3. **Create named ranges** for easier formula writing:
   ```excel
   # Create a dynamic named range for the matrix
   =SimilarityMatrix!$C$2:INDEX(SimilarityMatrix!$1:$1048576, COUNTA(SimilarityMatrix!$A:$A), COUNTA(SimilarityMatrix!$1:$1))
   ```

4. **Save frequently** - large similarity matrices can be resource-intensive

5. **Consider 64-bit Excel** for very large datasets

## Common Use Cases

### Finding Optimal Threshold:
1. Look at the Threshold Summary table
2. Find the "elbow point" where groups found drops significantly
3. Balance between too many groups (low threshold) and too few (high threshold)

### Analyzing Specific Products:
1. Use VLOOKUP to find a product's row in the similarity matrix
2. Sort by similarity score to find its closest matches
3. Adjust threshold to see how its group membership changes

### Quality Assurance:
1. Randomly sample groups at different thresholds
2. Verify that grouped products are truly similar
3. Adjust threshold based on business requirements

## Troubleshooting

**Excel becomes slow**: 
- Work with a sample of the data
- Use Excel calculations set to Manual
- Consider Power Pivot for better performance

**Formulas showing #REF!**:
- Check that the similarity matrix range is correct
- Ensure product names match exactly

**PivotTable not updating**:
- Right-click → **Refresh**
- Check data source connections

## Video Tutorial Links

For visual learners, search for these topics on YouTube:
- "Excel Power Pivot tutorial"
- "Dynamic threshold analysis in Excel"
- "Create heatmap from similarity matrix Excel"
