import pandas as pd
import numpy as np
from io import BytesIO

from .product_grouping import create_grouped_results, get_group_analyses


def compute_threshold_explorer(
    similarity_matrix,
    product_names,
    product_df,
    selected_output_columns,
    min_group_size,
    max_groups,
    threshold_range,
    conservative_grouping=True,
):
    start_thr, end_thr = threshold_range
    threshold_values = list(range(start_thr, end_thr + 1, 5))

    summary_rows = []
    frames = []

    for thr in threshold_values:
        analyses_thr = get_group_analyses(
            similarity_matrix=similarity_matrix,
            product_names=product_names,
            similarity_threshold=thr,
            min_group_size=min_group_size,
            max_groups=max_groups,
            conservative_grouping=conservative_grouping,
        )

        covered = int(sum(a["group_size"] for a in analyses_thr))
        largest = int(max((a["group_size"] for a in analyses_thr), default=0))
        avg_group_similarity = float(np.mean([a["avg_similarity"] for a in analyses_thr])) if analyses_thr else 0.0

        summary_rows.append(
            {
                "Threshold": thr,
                "Groups Found": len(analyses_thr),
                "Products in Groups": covered,
                "Largest Group Size": largest,
                "Avg Group Similarity": round(avg_group_similarity, 2),
            }
        )

        if analyses_thr:
            frame = create_grouped_results(analyses_thr, product_df, selected_output_columns)
            frame.insert(0, "Threshold", thr)
            frames.append(frame)

    summary_df = pd.DataFrame(summary_rows)
    rows_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return threshold_values, summary_df, rows_df


def build_threshold_workbook(group_rows_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        group_rows_df.to_excel(writer, index=False, sheet_name="Threshold Groups")
        summary_df.to_excel(writer, index=False, sheet_name="Threshold Summary")
    return output.getvalue()


def build_enhanced_threshold_workbook(
    similarity_matrix,
    product_names,
    product_df,
    group_rows_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    selected_output_columns=None
) -> bytes:
    """
    Build an enhanced Excel workbook with similarity matrix data for Excel-based threshold exploration.
    
    This workbook includes:
    1. Similarity matrix (sampled for large datasets)
    2. Product list with all attributes
    3. Threshold summary data
    4. Group data at all thresholds
    5. Instructions for Excel analysis
    """
    if selected_output_columns is None:
        selected_output_columns = []
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # 1. Similarity Matrix (sample if too large)
        similarity_sheet = workbook.add_worksheet('Similarity Matrix')
        
        # Add headers for similarity matrix
        similarity_sheet.write(0, 0, 'Product Index')
        similarity_sheet.write(0, 1, 'Product Name')
        
        # Determine sampling
        max_products = 1000  # Limit for Excel display
        n_products = len(product_names)
        
        if n_products <= max_products:
            # Show all products
            indices = list(range(n_products))
            similarity_sheet.write_row(0, 2, product_names)
        else:
            # Sample products
            import random
            random.seed(42)
            indices = random.sample(range(n_products), max_products)
            sampled_names = [product_names[i] for i in indices]
            similarity_sheet.write_row(0, 2, sampled_names)
        
        # Write similarity data
        for i, idx in enumerate(indices):
            similarity_sheet.write(i + 1, 0, idx)
            similarity_sheet.write(i + 1, 1, product_names[idx])
            row_data = similarity_matrix[idx, indices].tolist() if n_products <= max_products else similarity_matrix[idx, indices].tolist()
            similarity_sheet.write_row(i + 1, 2, row_data)
        
        # 2. Product List
        product_list_df = product_df.copy()
        product_list_df.insert(0, 'Product Index', range(len(product_df)))
        product_list_df.to_excel(writer, index=False, sheet_name='Product List', startrow=0)
        
        # 3. Threshold Summary
        summary_df.to_excel(writer, index=False, sheet_name='Threshold Summary')
        
        # 4. Threshold Groups
        group_rows_df.to_excel(writer, index=False, sheet_name='Threshold Groups')
        
        # 5. Instructions Sheet
        instructions_sheet = workbook.add_worksheet('Excel Instructions')
        
        # Add formatting
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#E7E6E6',
            'border': 1
        })
        
        # Write instructions
        instructions_sheet.set_column(0, 0, 80)
        instructions_sheet.set_column(1, 1, 20)
        
        instructions_sheet.write(0, 0, 'Creating a Threshold Explorer in Excel', header_format)
        instructions_sheet.write(1, 0, '')
        
        # Method 1: Using Power Pivot (Recommended)
        instructions_sheet.write(2, 0, 'Method 1: Using Power Pivot (Recommended)', title_format)
        instructions_sheet.write(3, 0, '')
        instructions_sheet.write(4, 0, '1. Go to Power Pivot > Add to Data Model')
        instructions_sheet.write(5, 0, '2. Add both "Product List" and "Similarity Matrix" tables to the model')
        instructions_sheet.write(6, 0, '3. Create a relationship between Product Index columns')
        instructions_sheet.write(7, 0, '4. Add a calculated column: Similarity % = [Similarity Value] * 100')
        instructions_sheet.write(8, 0, '5. Create a measure: Count Above Threshold = COUNTROWS(FILTER(SimilarityMatrix, [Similarity %] > SELECTEDVALUE(ThresholdSummary[Threshold])))')
        instructions_sheet.write(9, 0, '6. Add a PivotTable with Threshold as slicer')
        instructions_sheet.write(10, 0, '')
        
        # Method 2: Using Formulas
        instructions_sheet.write(11, 0, 'Method 2: Using Excel Formulas', title_format)
        instructions_sheet.write(12, 0, '')
        instructions_sheet.write(13, 0, '1. Create a new sheet for analysis')
        instructions_sheet.write(14, 0, '2. Copy product names from Product List to column A')
        instructions_sheet.write(15, 0, '3. In cell B1, enter your desired threshold (e.g., 75)')
        instructions_sheet.write(16, 0, '4. In column B, use formula to count matches: =COUNTIF(SimilarityMatrix!2:1000,">"&$B$1/100)')
        instructions_sheet.write(17, 0, '5. Use conditional formatting to highlight cells above threshold')
        instructions_sheet.write(18, 0, '')
        
        # Method 3: Using Pivot Tables
        instructions_sheet.write(19, 0, 'Method 3: Using Pivot Tables', title_format)
        instructions_sheet.write(20, 0, '')
        instructions_sheet.write(21, 0, '1. From Similarity Matrix, create a PivotTable')
        instructions_sheet.write(22, 0, '2. Add Product Index to Rows')
        instructions_sheet.write(23, 0, '3. Add VALUES to Values area (as Average)')
        instructions_sheet.write(24, 0, '4. Add a slicer for threshold using Value Filters')
        instructions_sheet.write(25, 0, '5. Group results to create clusters')
        instructions_sheet.write(26, 0, '')
        
        # Additional Tips
        instructions_sheet.write(27, 0, 'Additional Tips:', title_format)
        instructions_sheet.write(28, 0, '')
        instructions_sheet.write(29, 0, '• Use conditional formatting with color scales to visualize similarity patterns')
        instructions_sheet.write(30, 0, '• Create a dynamic named range for threshold: =INDIRECT("SimilarityMatrix!R2C2:R"&COUNTA(SimilarityMatrix!A:A)&"C"&COUNTA(SimilarityMatrix!1:1))')
        instructions_sheet.write(31, 0, '• For large datasets, use Power Query to load and transform the data')
        instructions_sheet.write(32, 0, '• Create a dashboard with charts showing groups vs threshold')
        
    return output.getvalue()
