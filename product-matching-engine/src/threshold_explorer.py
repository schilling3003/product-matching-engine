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
    Build a self-contained Excel workbook for threshold exploration.
    
    Uses native Excel features (Data Validation dropdown, VLOOKUP formulas, charts)
    with pre-computed group data at multiple thresholds. No VBA, no Power Pivot required.
    
    This workbook includes:
    1. Dashboard with threshold dropdown selector and auto-updating metrics
    2. Summary sheet with statistics for each threshold
    3. Groups sheet with all group membership data
    4. Optional similarity heatmap for visualization (small datasets only)
    """
    from .excel_export import build_threshold_explorer_workbook
    
    return build_threshold_explorer_workbook(
        summary_df=summary_df,
        groups_df=group_rows_df,
        similarity_matrix=similarity_matrix,
        product_names=product_names,
        max_heatmap_products=500
    )
