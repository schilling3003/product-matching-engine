# Threshold Explorer Manual Setup (Excel + Power Pivot)

Use this guide when VBA automation fails for any step (especially relationship creation).

## Prerequisites

1. Open the exported workbook (`threshold_explorer_analysis.xlsx`) in **desktop Excel**.
2. Ensure the **Power Pivot** add-in is enabled:
   - File -> Options -> Add-ins
   - Manage: COM Add-ins -> Go
   - Check **Microsoft Power Pivot for Excel**
3. Save workbook as a macro-enabled file if needed (`.xlsm`) before further edits.

---

## Step 1) Confirm required sheets and headers

Verify these sheets exist:
- `Product List`
- `Similarity Matrix`
- `Threshold Summary` (optional)

Verify header row names:
- `Product List` contains `Product Index`
- `Similarity Matrix` contains `Product Index`

> Header names must match exactly (including spacing/case) for easiest setup.

---

## Step 2) Convert ranges to Excel Tables

For each sheet below:

### Product List
1. Go to `Product List`.
2. Click any cell in the data.
3. Press **Ctrl+T** (My table has headers = checked).
4. Rename table on Table Design tab to: `ProductList`.

### Similarity Matrix
1. Go to `Similarity Matrix`.
2. Click any cell in the data.
3. Press **Ctrl+T** (My table has headers = checked).
4. Rename table to: `SimilarityMatrix`.

### Threshold Summary (optional)
1. Go to `Threshold Summary`.
2. Press **Ctrl+T** and rename to: `ThresholdSummary`.

---

## Step 3) Add tables to the Data Model manually

For each table (`ProductList`, `SimilarityMatrix`, optionally `ThresholdSummary`):

1. Click inside the table.
2. Go to **Power Pivot** tab.
3. Click **Add to Data Model**.
4. Wait until Power Pivot finishes loading.

Validation:
- Open **Power Pivot -> Manage**.
- Confirm each table appears as a tab in the Power Pivot window.

---

## Step 4) Create relationship manually

In **Power Pivot -> Manage**:

1. Go to Diagram View (or Design -> Create Relationship).
2. Create this relationship:
   - From table: `SimilarityMatrix`
   - Column: `Product Index`
   - To table: `ProductList`
   - Column: `Product Index`
3. Save.

If table names in Power Pivot include spaces (e.g., `Similarity Matrix`), use those exact displayed names.

---

## Step 5) Calculated columns/measures (optional)

For the current exported workbook layout, `SimilarityMatrix` is a wide matrix and usually **does not** have a single column named `Similarity Value`.

Because of that, skip the previous `Similarity %` DAX step unless you first create a normalized/pairwise table (one row per product pair with one similarity value column).

### What to do now (recommended)

- Skip DAX calculated column/measure creation.
- Continue with Step 6 worksheet formulas for threshold analysis.

### Optional (only if you create a normalized table)

If you manually build a normalized table containing a numeric column `Similarity Value`, then you can add:

```DAX
Similarity % := [Similarity Value] * 100
```

and:

```DAX
Average Similarity := AVERAGE('SimilarityPairs'[Similarity %])
```

Replace `SimilarityPairs` with your actual normalized table name.

---

## Step 6) Build analysis sheet manually (if needed)

In `Analysis` sheet:

- `A3`: Enter label `Enter Threshold (%):`
- `B3`: Enter threshold value (e.g., `75`)
- `A5`: Label `Products Above Threshold:`
- `B5` formula:

```excel
=COUNTIF('Similarity Matrix'!C:Z,">" & B3/100)
```

- `A6`: Label `Potential Groups:`
- `B6` formula:

```excel
=ROUND(B5/2,0)
```

---

## Step 7) Troubleshooting checklist

### Relationship errors
- Confirm both model tables contain `Product Index`.
- Ensure column data types match in both tables (e.g., both whole number or both text).
- Remove and recreate relationship if necessary.

### "Column not found" in Power Pivot
- Recheck table headers in worksheet source.
- Refresh Data Model after header corrections.
- If error mentions `Similarity Value`, your model likely has a wide matrix table; skip DAX step unless you create a normalized similarity table.

### Data model load failures
- Re-add table via **Power Pivot -> Add to Data Model**.
- Ensure workbook is saved locally (not cloud-locked/temp path).
- Close/reopen Excel and retry.

---

## Suggested manual run order when VBA partially fails

1. Run VBA step for validation/sheet creation only.
2. Perform Steps 2-4 above manually.
3. Skip Step 5 unless you created a normalized similarity table.
4. Run VBA steps for analysis/dashboard only.
5. If still blocked, complete analysis formulas manually.

---

## Notes for this project

- If relationship creation fails in VBA, manual relationship setup is acceptable.
- After manual setup is confirmed working, we can simplify VBA to skip only the failing step and keep the rest automated.
