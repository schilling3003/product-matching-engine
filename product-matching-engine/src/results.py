import heapq
import pandas as pd
import numpy as np


def _sanitize_for_streamlit(df: pd.DataFrame) -> pd.DataFrame:
    """Return an Arrow-safe DataFrame for Streamlit display/export."""
    safe_df = df.copy()
    for col in safe_df.columns:
        if safe_df[col].dtype == "object":
            safe_df[col] = safe_df[col].apply(
                lambda x: str(x) if isinstance(x, (dict, list, set, tuple)) else x
            )
    return safe_df


def _get_process_memory_mb():
    """Best-effort process memory in MB for lightweight diagnostics."""
    try:
        import os
        import psutil

        return psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024
    except Exception:
        return None


def _limit_between_file_streaming_results(streaming_results, max_matches_per_product, progress_callback=None):
    """Keep only top-k matches per customer from streaming tuple results."""
    if not streaming_results or max_matches_per_product <= 0:
        return streaming_results

    top_k_heaps = {}
    total = len(streaming_results)

    for idx, rec in enumerate(streaming_results):
        if progress_callback is not None and (idx % 5000 == 0 or idx == total - 1):
            progress_callback((idx + 1) / total, idx + 1, total)

        customer_idx = int(rec[0])
        score = float(rec[2])
        heap = top_k_heaps.setdefault(customer_idx, [])

        if len(heap) < max_matches_per_product:
            heapq.heappush(heap, (score, rec))
        elif score > heap[0][0]:
            heapq.heapreplace(heap, (score, rec))

    limited_results = []
    for heap in top_k_heaps.values():
        limited_results.extend(rec for _, rec in sorted(heap, key=lambda x: x[0], reverse=True))

    return limited_results


def update_results_with_additional_columns(match_results, new_column_config, is_within_file):
    """Update existing results with additional columns without re-running matching."""
    if not match_results or "raw_matches" not in match_results:
        return match_results["results_df"]

    raw_matches = match_results["raw_matches"]
    customer_df = match_results["customer_df"]
    catalog_df = match_results["catalog_df"]
    settings = match_results.get("settings", {})

    customer_display_col = match_results["column_config"]["customer"]["product_cols"][0]
    catalog_display_col = match_results["column_config"]["catalog"]["product_cols"][0]

    # Get new output columns
    new_customer_out_cols = new_column_config.get("customer", {}).get("output_cols", [])
    new_catalog_out_cols = new_column_config.get("catalog", {}).get("output_cols", [])

    # Build new results DataFrame
    rows = []
    for idx, rec in enumerate(raw_matches):
        # Handle both old format (6 elements) and new format (7 elements with size)
        if len(rec) == 7:
            i, j, combined, tfidf_s, fuzzy_s, gtin_s, size_s = rec
        else:
            i, j, combined, tfidf_s, fuzzy_s, gtin_s = rec
            size_s = 0

        if is_within_file:
            entry = {
                "Product 1": customer_df.iloc[i][customer_display_col],
                "Product 2": catalog_df.iloc[j][catalog_display_col],
                "Confidence Score": f"{combined:.2f}%",
                "TF-IDF Score": f"{tfidf_s:.2f}%",
                "Fuzzy Score": f"{fuzzy_s:.2f}%",
            }
        else:
            entry = {
                "Customer Product": customer_df.iloc[i][customer_display_col],
                "Catalog Product": catalog_df.iloc[j][catalog_display_col],
                "Confidence Score": f"{combined:.2f}%",
                "TF-IDF Score": f"{tfidf_s:.2f}%",
                "Fuzzy Score": f"{fuzzy_s:.2f}%",
            }

        # Add GTIN details if available
        if gtin_s > 0:
            entry["GTIN Score"] = f"{gtin_s:.2f}%"

        # Add size details if available
        if settings.get("size_weight", 0) > 0 and size_s > 0:
            entry["Size Score"] = f"{size_s:.2f}%"
            customer_size = customer_df.iloc[i].get("standardized_size", "")
            catalog_size = catalog_df.iloc[j].get("standardized_size", "")
            if customer_size and catalog_size:
                entry[f"Product 1 Size" if is_within_file else "Customer Size"] = customer_size
                entry[f"Product 2 Size" if is_within_file else "Catalog Size"] = catalog_size

        # Add new customer columns
        for col in new_customer_out_cols:
            if col in customer_df.columns:
                entry[f"Product 1 {col}" if is_within_file else f"Customer {col}"] = customer_df.iloc[i][col]

        # Add new catalog columns
        for col in new_catalog_out_cols:
            if col in catalog_df.columns:
                entry[f"Product 2 {col}" if is_within_file else f"Catalog {col}"] = catalog_df.iloc[j][col]

        rows.append(entry)

    return pd.DataFrame(rows)


def convert_streaming_results_to_dataframe(
    streaming_results,
    cleaned_customer_df,
    cleaned_catalog_df,
    column_config,
    is_within_file,
    settings,
    gtin_details=None,
    progress_callback=None,
):
    """Convert chunked extraction results (list of tuples) to the expected DataFrame format."""
    if not streaming_results:
        return pd.DataFrame()

    customer_display_col = column_config["customer"]["product_cols"][0]
    catalog_display_col = column_config["catalog"]["product_cols"][0]
    customer_out_cols = column_config["customer"].get("output_cols", [])
    catalog_out_cols = column_config["catalog"].get("output_cols", [])

    # Pre-fetch numpy arrays for extremely fast and memory-efficient lookups
    cust_disp_vals = cleaned_customer_df[customer_display_col].values
    cat_disp_vals = cleaned_catalog_df[catalog_display_col].values

    cust_size_vals = cleaned_customer_df.get(
        "standardized_size", pd.Series([""] * len(cleaned_customer_df))
    ).values
    cat_size_vals = cleaned_catalog_df.get(
        "standardized_size", pd.Series([""] * len(cleaned_catalog_df))
    ).values

    cust_out_vals = {
        col: cleaned_customer_df[col].values for col in customer_out_cols if col in cleaned_customer_df.columns
    }
    cat_out_vals = {
        col: cleaned_catalog_df[col].values for col in catalog_out_cols if col in cleaned_catalog_df.columns
    }

    rows = []
    total_records = len(streaming_results)
    for idx, rec in enumerate(streaming_results):
        if progress_callback is not None and (idx % 5000 == 0 or idx == total_records - 1):
            progress_callback((idx + 1) / total_records, idx + 1, total_records)

        # Handle both old format (6 elements) and new format (7 elements with size)
        if len(rec) == 7:
            i, j, combined, tfidf_s, fuzzy_s, gtin_s, size_s = rec
        else:
            i, j, combined, tfidf_s, fuzzy_s, gtin_s = rec
            size_s = 0

        if is_within_file:
            entry = {
                "Product 1": cust_disp_vals[i],
                "Product 2": cat_disp_vals[j],
                "Confidence Score": f"{combined:.2f}%",
                "TF-IDF Score": f"{tfidf_s:.2f}%",
                "Fuzzy Score": f"{fuzzy_s:.2f}%",
            }
        else:
            entry = {
                "Customer Product": cust_disp_vals[i],
                "Catalog Product": cat_disp_vals[j],
                "Confidence Score": f"{combined:.2f}%",
                "TF-IDF Score": f"{tfidf_s:.2f}%",
                "Fuzzy Score": f"{fuzzy_s:.2f}%",
            }

        if gtin_s > 0:
            entry["GTIN Score"] = f"{gtin_s:.2f}%"
        if gtin_details and (i, j) in gtin_details:
            d = gtin_details[(i, j)]
            entry["GTIN Match Type"] = d["match_type"]
            entry["Matching GTINs"] = ", ".join(d["matching_gtins"][:3])

        # Add size match details if size matching is enabled
        if settings.get("size_weight", 0) > 0 and size_s > 0:
            entry["Size Score"] = f"{size_s:.2f}%"
            # Add standardized sizes for reference
            customer_size = cust_size_vals[i]
            catalog_size = cat_size_vals[j]
            if customer_size and catalog_size:
                entry[f"Product 1 Size" if is_within_file else "Customer Size"] = customer_size
                entry[f"Product 2 Size" if is_within_file else "Catalog Size"] = catalog_size

        for col, vals in cust_out_vals.items():
            entry[f"Product 1 {col}" if is_within_file else f"Customer {col}"] = vals[i]
        for col, vals in cat_out_vals.items():
            entry[f"Product 2 {col}" if is_within_file else f"Catalog {col}"] = vals[j]

        rows.append(entry)

    return pd.DataFrame(rows)
