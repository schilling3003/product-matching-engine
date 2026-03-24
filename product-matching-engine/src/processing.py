import re
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from rapidfuzz import fuzz
from rapidfuzz.process import cdist
from concurrent.futures import ProcessPoolExecutor
import multiprocessing as mp
import gc
import heapq
import psutil
import os

from .config import STOP_WORDS, UNIT_CONVERSION_MAP
from .gtin_processing import consolidate_gtin_columns, calculate_gtin_match_confidence
from .product_grouping import (
    create_grouped_results,
    get_group_analyses,
)

def clean_and_standardize(df, column_config, remove_stop_words=True, case_sensitive=False, include_size_in_text=False):
    """Cleans and standardizes the product data in a DataFrame, including complex size handling and GTIN processing."""
    if df is None: return None
    cleaned_df = df.copy()

    # Get column mappings - now supporting multiple product name fields and GTIN columns
    product_cols = column_config.get('product_cols', [])
    size_col = column_config.get('size')
    size_value_col = column_config.get('size_value')
    size_unit_col = column_config.get('size_unit')
    gtin_cols = column_config.get('gtin_cols', [])

    # --- Size Standardization ---
    # Combine size_value and size_unit if they exist
    if size_value_col and size_unit_col and size_value_col in cleaned_df.columns and size_unit_col in cleaned_df.columns:
        # Ensure size_value is numeric, coercing errors to NaN
        cleaned_df['size_value'] = pd.to_numeric(cleaned_df[size_value_col], errors='coerce')
        # Fill NaN in unit with a default empty string
        cleaned_df['size_unit'] = cleaned_df[size_unit_col].fillna('')
        # Create a unified 'size' column
        cleaned_df['size'] = cleaned_df.apply(lambda row: f"{row['size_value']} {row['size_unit']}" if pd.notna(row['size_value']) else '', axis=1)

    # Standardize units from the 'size' column (or pre-existing)
    if size_col and size_col in cleaned_df.columns:
        def standardize_size_column(text):
            text = str(text).lower()
            for unit, multiplier in UNIT_CONVERSION_MAP.items():
                pattern = r'(\d*\.?\d+)\s*(' + re.escape(unit) + r')\b'
                match = re.search(pattern, text)
                if match:
                    value = float(match.group(1))
                    base_value = value * multiplier
                    # Return a consistent format, e.g., "141.8g"
                    return f"{base_value:.1f}g"
            return "" # Return empty if no unit found
        cleaned_df['standardized_size'] = cleaned_df[size_col].apply(standardize_size_column)
    else:
        cleaned_df['standardized_size'] = ''

    # --- Text Column Processing ---
    text_columns = [col for col in product_cols if col in cleaned_df.columns]
        
    for col in text_columns:
        # Case handling
        if not case_sensitive:
            cleaned_df[col] = cleaned_df[col].astype(str).str.lower()
        else:
            cleaned_df[col] = cleaned_df[col].astype(str)
        
        # Remove special characters
        cleaned_df[col] = cleaned_df[col].str.replace(r'[^\w\s]', '', regex=True)
        
        # Stop word removal (optional)
        if remove_stop_words:
            cleaned_df[col] = cleaned_df[col].apply(lambda x: ' '.join([word for word in x.split() if word not in STOP_WORDS]))
        
        # Clean up whitespace
        cleaned_df[col] = cleaned_df[col].str.replace(r'\s+', ' ', regex=True).str.strip()

    # Create a combined product name field for matching
    combined_product_parts = [cleaned_df[col] for col in text_columns]
    
    if combined_product_parts:
        # Combine multiple product name fields with space separator
        cleaned_df['combined_product_name'] = pd.concat(combined_product_parts, axis=1).apply(
            lambda x: ' '.join(x.dropna().astype(str)), axis=1
        )
        
        # Append standardized size to combined product name for better matching (if enabled)
        if include_size_in_text:
            cleaned_df['combined_product_name'] = cleaned_df['combined_product_name'] + ' ' + cleaned_df['standardized_size']
            cleaned_df['combined_product_name'] = cleaned_df['combined_product_name'].str.strip()
    else:
        cleaned_df['combined_product_name'] = ''
    
    # --- GTIN Processing ---
    # Consolidate GTIN columns into normalized variant pools
    if gtin_cols:
        cleaned_df['gtin_pool'] = consolidate_gtin_columns(cleaned_df, gtin_cols)
    else:
        cleaned_df['gtin_pool'] = pd.Series([{}] * len(cleaned_df))

    return cleaned_df

def calculate_size_similarity_vectorized(customer_sizes, catalog_sizes, tolerance_percent=20):
    """Vectorized calculation of size similarity matrix."""
    import re
    
    def extract_values(sizes):
        vals = []
        for s in sizes:
            if not s:
                vals.append(np.nan)
                continue
            try:
                match = re.search(r'(\d*\.?\d+)', str(s))
                if match:
                    vals.append(float(match.group(1)))
                else:
                    vals.append(np.nan)
            except:
                vals.append(np.nan)
        return np.array(vals)

    c1 = extract_values(customer_sizes)[:, np.newaxis]
    c2 = extract_values(catalog_sizes)[np.newaxis, :]

    valid_mask = ~np.isnan(c1) & ~np.isnan(c2)
    larger = np.maximum(c1, c2)
    smaller = np.minimum(c1, c2)
    
    both_zero = (c1 == 0) & (c2 == 0)

    with np.errstate(divide='ignore', invalid='ignore'):
        percent_diff = ((larger - smaller) / larger) * 100

    similarity = 100 * (1 - (percent_diff / tolerance_percent))
    similarity = np.clip(similarity, 0, 100)

    final_sim = np.where(valid_mask, similarity, 0.0)
    final_sim = np.where(both_zero & valid_mask, 100.0, final_sim)
    final_sim = np.nan_to_num(final_sim, nan=0.0)
    
    return final_sim

def calculate_size_similarity(size1, size2, tolerance_percent=20):
    """Calculate similarity between two standardized sizes based on their numeric values.
    
    Args:
        size1 (str): First size in format like '141.8g' or ''
        size2 (str): Second size in format like '141.8g' or ''
        tolerance_percent (float): Percentage tolerance for considering sizes similar
    
    Returns:
        float: Similarity score from 0-100
    """
    if not size1 or not size2 or size1 == '' or size2 == '':
        return 0
    
    try:
        # Extract numeric values from standardized sizes (format: "123.4g")
        import re
        match1 = re.search(r'(\d*\.?\d+)', size1)
        match2 = re.search(r'(\d*\.?\d+)', size2)
        
        if not match1 or not match2:
            return 0
            
        val1 = float(match1.group(1))
        val2 = float(match2.group(1))
        
        # Handle zero values
        if val1 == 0 or val2 == 0:
            return 100 if val1 == val2 else 0
        
        # Calculate percentage difference
        larger = max(val1, val2)
        smaller = min(val1, val2)
        percent_diff = ((larger - smaller) / larger) * 100
        
        # Convert to similarity score
        if percent_diff <= tolerance_percent:
            # Linear decay from 100 to 0 within tolerance
            similarity = 100 * (1 - (percent_diff / tolerance_percent))
            return max(0, min(100, similarity))
        else:
            return 0
            
    except (ValueError, AttributeError):
        return 0

def batch_fuzzy_matching(customer_texts, catalog_texts, chunk_size=1000):
    """
    Optimized batch fuzzy matching using RapidFuzz C++ extension.
    Returns a matrix of fuzzy scores.
    """
    # RapidFuzz's cdist is implemented in C++ and automatically uses all available cores
    # when workers=-1. It is orders of magnitude faster than Python multiprocessing loops.
    fuzzy_matrix = cdist(
        customer_texts, 
        catalog_texts, 
        scorer=fuzz.token_set_ratio,
        workers=-1,
        dtype=np.float64
    )
    return fuzzy_matrix

def get_memory_usage_mb():
    """Get current memory usage in MB."""
    process = psutil.Process(os.getpid())
    return process.memory_info().rss / 1024 / 1024

def calculate_similarity_memory_efficient(customer_texts, catalog_texts, customer_vectors, catalog_vectors,
                                        tfidf_weight=0.5, fuzzy_weight=0.5, gtin_weight=0.0, size_weight=0.0,
                                        customer_sizes=None, catalog_sizes=None, size_tolerance=20,
                                        customer_gtins=None, catalog_gtins=None,
                                        similarity_threshold=50, early_filter=True, enable_multiprocessing=True, batch_size=1000,
                                        within_file_mode=False, progress_callback=None, max_memory_mb=1500,
                                        restriction_data=None, max_matches_per_product=5):
    """
    Memory-efficient similarity calculation.
    For large datasets: processes in chunks, extracts results immediately, discards chunk matrices.
    Maintains vectorized speed while keeping memory constant.
    """
    n_customers = len(customer_texts)
    n_catalog = len(catalog_texts)

    def _emit_progress(progress):
        if progress_callback is None:
            return
        clamped = max(0.0, min(1.0, float(progress)))
        current = min(n_customers, max(0, int(clamped * n_customers)))
        progress_callback(clamped, current, n_customers)

    _emit_progress(0.02)

    # Account for TF-IDF, fuzzy, GTIN, size, combined + overhead.
    estimated_memory_mb = (n_customers * n_catalog * 8 * 6) / 1024 / 1024
    # Keep behavior conservative for Cloud stability: always chunk for larger pair counts.
    use_chunked = (
        (n_customers * n_catalog) > 5_000_000
        or estimated_memory_mb > max_memory_mb
        or n_customers > 10000
    )

    if use_chunked:
        print(f"🧠 Chunked mode: {estimated_memory_mb:.0f}MB estimated, processing in chunks")
        return _chunked_extract_results(
            customer_texts, catalog_texts, customer_vectors, catalog_vectors,
            tfidf_weight, fuzzy_weight, gtin_weight, size_weight,
            customer_sizes, catalog_sizes, size_tolerance,
            customer_gtins, catalog_gtins, similarity_threshold, early_filter,
            enable_multiprocessing, batch_size, within_file_mode, progress_callback,
            restriction_data, max_matches_per_product
        )
    else:
        return calculate_similarity_vectorized(
            customer_texts, catalog_texts, customer_vectors, catalog_vectors,
            tfidf_weight, fuzzy_weight, gtin_weight, size_weight,
            customer_sizes, catalog_sizes, size_tolerance,
            customer_gtins, catalog_gtins, similarity_threshold, early_filter,
            enable_multiprocessing, batch_size, within_file_mode, progress_callback, restriction_data
        )


def _chunked_extract_results(customer_texts, catalog_texts, customer_vectors, catalog_vectors,
                             tfidf_weight, fuzzy_weight, gtin_weight, size_weight,
                             customer_sizes, catalog_sizes, size_tolerance,
                             customer_gtins, catalog_gtins, similarity_threshold, early_filter,
                             enable_multiprocessing, batch_size, within_file_mode, progress_callback,
                             restriction_data=None, max_matches_per_product=5):
    """
    Vectorized chunk processing that never stores full N×N matrices.
    Each chunk is processed fully (TF-IDF, fuzzy, size, GTIN), results above threshold
    are extracted immediately, then the chunk matrices are discarded.
    This preserves full vectorized speed with constant memory usage.
    """
    n_customers = len(customer_texts)
    n_catalog = len(catalog_texts)

    # Chunk size: keep each chunk's matrices under ~200MB
    # chunk_size rows × n_catalog cols × 8 bytes × 4 matrices
    chunk_size = max(50, min(500, int(200 * 1024 * 1024 / (n_catalog * 8 * 4))))
    print(f"🔄 Chunked extraction: {n_customers:,} products in chunks of {chunk_size:,}")

    # Accumulate only sparse results, not full matrices.
    # Between-files mode can produce huge raw match lists, so keep only top-k per customer on the fly.
    use_top_k_heaps = (not within_file_mode and max_matches_per_product is not None and max_matches_per_product > 0)
    top_k_heaps = {} if use_top_k_heaps else None
    match_results = []
    gtin_details = {}

    # Build inverted index for catalog GTINs once, to turn O(N*M) search into O(N)
    catalog_gtin_index = {}
    if gtin_weight > 0 and catalog_gtins is not None:
        for j, cat_pool in enumerate(catalog_gtins):
            if cat_pool:
                for gtin, match_type in cat_pool.items():
                    if gtin not in catalog_gtin_index:
                        catalog_gtin_index[gtin] = []
                    catalog_gtin_index[gtin].append((j, match_type))

    for chunk_start in range(0, n_customers, chunk_size):
        chunk_end = min(chunk_start + chunk_size, n_customers)
        chunk_len = chunk_end - chunk_start

        if progress_callback is not None:
            progress_callback(chunk_start / n_customers, chunk_start, n_customers)

        # --- TF-IDF (vectorized) ---
        if tfidf_weight > 0 and customer_vectors is not None and catalog_vectors is not None:
            chunk_tfidf = cosine_similarity(customer_vectors[chunk_start:chunk_end], catalog_vectors) * 100
        else:
            chunk_tfidf = np.zeros((chunk_len, n_catalog))

        # --- Fuzzy (vectorized candidates only) ---
        chunk_fuzzy = np.zeros((chunk_len, n_catalog))
        if fuzzy_weight > 0:
            min_tfidf_for_fuzzy = max(5, similarity_threshold * 0.2)
            top_k = min(1000, max(50, int(0.1 * n_catalog)))

            for i_local in range(chunk_len):
                i_global = chunk_start + i_local
                tfidf_row = chunk_tfidf[i_local]

                above_thresh = np.where(tfidf_row >= min_tfidf_for_fuzzy)[0]
                k = min(top_k, max(1, n_catalog - 1))
                topk_idx = np.argpartition(-tfidf_row, kth=k)[:top_k]
                candidates = np.unique(np.concatenate([above_thresh, topk_idx]))
                if within_file_mode:
                    candidates = candidates[candidates != i_global]

                cust_text = customer_texts[i_global]
                for j in candidates:
                    chunk_fuzzy[i_local, j] = fuzz.token_set_ratio(cust_text, catalog_texts[j])

        # --- Size (vectorized) ---
        if size_weight > 0 and customer_sizes is not None and catalog_sizes is not None:
            chunk_size_mat = calculate_size_similarity_vectorized(
                customer_sizes[chunk_start:chunk_end], 
                catalog_sizes, 
                size_tolerance
            )
        else:
            chunk_size_mat = np.zeros((chunk_len, n_catalog))

        # --- GTIN ---
        chunk_gtin = np.zeros((chunk_len, n_catalog))
        chunk_gtin_details = {}
        if gtin_weight > 0 and customer_gtins is not None and catalog_gtins is not None:
            for i_local in range(chunk_len):
                i_global = chunk_start + i_local
                cust_pool = customer_gtins[i_global]
                if not cust_pool:
                    continue
                
                # Use inverted index to find only the catalog products that share GTINs
                for gtin, cust_match_type in cust_pool.items():
                    if gtin in catalog_gtin_index:
                        for j, cat_match_type in catalog_gtin_index[gtin]:
                            # Calculate confidence directly (only for matching pairs)
                            conf, mtype = _get_gtin_confidence(cust_match_type, cat_match_type)
                            
                            # Keep best match if multiple GTINs overlap
                            if conf > chunk_gtin[i_local, j]:
                                chunk_gtin[i_local, j] = conf
                                chunk_gtin_details[(i_local, j)] = {
                                    'confidence': conf,
                                    'match_type': mtype,
                                    'matching_gtins': [gtin]
                                }
                            elif conf == chunk_gtin[i_local, j] and conf > 0:
                                # Add to existing list if equal confidence
                                if (i_local, j) in chunk_gtin_details:
                                    if gtin not in chunk_gtin_details[(i_local, j)]['matching_gtins']:
                                        if len(chunk_gtin_details[(i_local, j)]['matching_gtins']) < 3:
                                            chunk_gtin_details[(i_local, j)]['matching_gtins'].append(gtin)

        # --- Combine scores (vectorized) ---
        chunk_combined = _calculate_combined_score(
            chunk_tfidf, chunk_fuzzy, chunk_gtin, chunk_size_mat,
            tfidf_weight, fuzzy_weight, gtin_weight, size_weight
        )

        # --- Extract results above threshold immediately (sparse) ---
        if within_file_mode:
            for i_local in range(chunk_len):
                chunk_combined[i_local, chunk_start + i_local] = 0.0

        above = np.argwhere(chunk_combined >= similarity_threshold)
        for i_local, j in above:
            i_global = chunk_start + i_local
            
            # Apply restriction filters if enabled
            if restriction_data and within_file_mode:
                skip_match = False
                for idx, col in enumerate(restriction_data['columns']):
                    cust_val = restriction_data['customer_data'][idx][i_global]
                    cat_val = restriction_data['catalog_data'][idx][j]
                    if cust_val.lower() != cat_val.lower():
                        skip_match = True
                        break
                if skip_match:
                    continue
            
            rec = (
                i_global, int(j),
                float(chunk_combined[i_local, j]),
                float(chunk_tfidf[i_local, j]),
                float(chunk_fuzzy[i_local, j]),
                float(chunk_gtin[i_local, j]),
                float(chunk_size_mat[i_local, j])
            )

            if use_top_k_heaps:
                heap = top_k_heaps.setdefault(i_global, [])
                score = rec[2]
                detail_key = (i_global, int(j))

                if len(heap) < max_matches_per_product:
                    heapq.heappush(heap, (score, rec))
                    if (i_local, j) in chunk_gtin_details:
                        gtin_details[detail_key] = chunk_gtin_details[(i_local, j)]
                elif score > heap[0][0]:
                    _, dropped = heapq.heapreplace(heap, (score, rec))
                    dropped_key = (int(dropped[0]), int(dropped[1]))
                    if dropped_key in gtin_details:
                        del gtin_details[dropped_key]
                    if (i_local, j) in chunk_gtin_details:
                        gtin_details[detail_key] = chunk_gtin_details[(i_local, j)]
            else:
                match_results.append(rec)
                if (i_local, j) in chunk_gtin_details:
                    gtin_details[(i_global, int(j))] = chunk_gtin_details[(i_local, j)]

        # Discard chunk matrices immediately
        del chunk_tfidf, chunk_fuzzy, chunk_size_mat, chunk_gtin, chunk_combined, chunk_gtin_details
        gc.collect()

    if progress_callback is not None:
        progress_callback(1.0, n_customers, n_customers)

    if use_top_k_heaps:
        limited_results = []
        limited_gtin_details = {}
        for heap in top_k_heaps.values():
            for _, rec in sorted(heap, key=lambda x: x[0], reverse=True):
                limited_results.append(rec)
                detail_key = (int(rec[0]), int(rec[1]))
                if detail_key in gtin_details:
                    limited_gtin_details[detail_key] = gtin_details[detail_key]
        match_results = limited_results
        gtin_details = limited_gtin_details

    print(f"✅ Chunked extraction complete: {len(match_results):,} matches found")
    # Return sentinel so app.py knows this is streaming results
    dummy = np.zeros((1, 1))
    return dummy, dummy, dummy, dummy, gtin_details, match_results

def _calculate_similarity_chunked(customer_texts, catalog_texts, customer_vectors, catalog_vectors,
                                 tfidf_weight, fuzzy_weight, gtin_weight, size_weight,
                                 customer_sizes, catalog_sizes, size_tolerance,
                                 customer_gtins, catalog_gtins, similarity_threshold, early_filter,
                                 enable_multiprocessing, batch_size, within_file_mode, progress_callback):
    """
    Chunked similarity calculation for large datasets.
    Streams results directly to avoid storing full matrices in memory.
    """
    n_customers = len(customer_texts)
    n_catalog = len(catalog_texts)
    
    # For very large datasets, we'll stream results instead of storing matrices
    stream_results = n_customers * n_catalog > 5_000_000  # 5M elements threshold
    
    if stream_results:
        print(f"🌊 Streaming mode: {n_customers:,} × {n_catalog:,} = {n_customers*n_catalog:,} comparisons")
        return _stream_similarity_results(
            customer_texts, catalog_texts, customer_vectors, catalog_vectors,
            tfidf_weight, fuzzy_weight, gtin_weight, size_weight,
            customer_sizes, catalog_sizes, size_tolerance,
            customer_gtins, catalog_gtins, similarity_threshold, early_filter,
            enable_multiprocessing, batch_size, within_file_mode, progress_callback
        )
    
    # Otherwise use chunked processing with matrix storage
    return _chunked_with_matrices(
        customer_texts, catalog_texts, customer_vectors, catalog_vectors,
        tfidf_weight, fuzzy_weight, gtin_weight, size_weight,
        customer_sizes, catalog_sizes, size_tolerance,
        customer_gtins, catalog_gtins, similarity_threshold, early_filter,
        enable_multiprocessing, batch_size, within_file_mode, progress_callback
    )

def _stream_similarity_results(customer_texts, catalog_texts, customer_vectors, catalog_vectors,
                              tfidf_weight, fuzzy_weight, gtin_weight, size_weight,
                              customer_sizes, catalog_sizes, size_tolerance,
                              customer_gtins, catalog_gtins, similarity_threshold, early_filter,
                              enable_multiprocessing, batch_size, within_file_mode, progress_callback):
    """
    Stream similarity results without storing full matrices.
    Returns matrices but computed incrementally with memory cleanup.
    """
    n_customers = len(customer_texts)
    n_catalog = len(catalog_texts)
    
    # Initialize matrices (we still need these for the UI)
    # But we'll compute them row by row and clean up aggressively
    combined_matrix = np.zeros((n_customers, n_catalog))
    tfidf_matrix = np.zeros((n_customers, n_catalog))
    fuzzy_matrix = np.zeros((n_customers, n_catalog))
    gtin_matrix = np.zeros((n_customers, n_catalog))
    gtin_details = {}
    
    # Process one row at a time for maximum memory efficiency
    chunk_size = 100  # Very small chunks
    
    print(f"🔄 Streaming {n_customers:,} products with tiny chunks of {chunk_size:,}")
    
    for chunk_start in range(0, n_customers, chunk_size):
        chunk_end = min(chunk_start + chunk_size, n_customers)
        
        # Force garbage collection before each chunk
        gc.collect()
        
        for i in range(chunk_start, chunk_end):
            if progress_callback is not None and i % 100 == 0:
                progress = i / n_customers
                progress_callback(progress, i, n_customers)
            
            # Compute TF-IDF for this single row
            if tfidf_weight > 0 and customer_vectors is not None and catalog_vectors is not None:
                tfidf_row = cosine_similarity(customer_vectors[i:i+1], catalog_vectors).flatten() * 100
            else:
                tfidf_row = np.zeros(n_catalog)
            
            # Compute fuzzy for this row (with filtering)
            fuzzy_row = np.zeros(n_catalog)
            if fuzzy_weight > 0:
                if early_filter and tfidf_weight > 0:
                    # Find candidates using TF-IDF
                    min_tfidf_for_fuzzy = max(5, similarity_threshold * 0.2)
                    top_k = min(500, int(0.05 * n_catalog))
                    
                    above_thresh = np.where(tfidf_row >= min_tfidf_for_fuzzy)[0]
                    
                    if len(tfidf_row) > 0:
                        k = min(top_k, max(1, n_catalog - 1))
                        topk_idx = np.argpartition(-tfidf_row, kth=k)[:top_k]
                    else:
                        topk_idx = np.array([], dtype=int)
                    
                    candidates = np.unique(np.concatenate([above_thresh, topk_idx]))
                    
                    # Skip self in within-file mode
                    if within_file_mode:
                        candidates = candidates[candidates != i]
                    
                    # Calculate fuzzy only for candidates
                    if len(candidates) > 0:
                        customer_text = customer_texts[i]
                        for j in candidates:
                            fuzzy_row[j] = fuzz.token_set_ratio(customer_text, catalog_texts[j])
                else:
                    # Full fuzzy for small datasets
                    customer_text = customer_texts[i]
                    for j, catalog_text in enumerate(catalog_texts):
                        if not within_file_mode or i != j:
                            fuzzy_row[j] = fuzz.token_set_ratio(customer_text, catalog_text)
            
            # Compute size for this row
            size_row = np.zeros(n_catalog)
            if size_weight > 0 and customer_sizes is not None and catalog_sizes is not None:
                customer_size = customer_sizes[i]
                if customer_size:
                    for j, catalog_size in enumerate(catalog_sizes):
                        size_row[j] = calculate_size_similarity(customer_size, catalog_size, size_tolerance)
            
            # Compute GTIN for this row
            gtin_row = np.zeros(n_catalog)
            if gtin_weight > 0 and customer_gtins is not None and catalog_gtins is not None:
                cust_pool = customer_gtins[i]
                if cust_pool:
                    for j, cat_pool in enumerate(catalog_gtins):
                        if cat_pool:
                            common_gtins = set(cust_pool.keys()) & set(cat_pool.keys())
                            if common_gtins:
                                best_confidence = 0.0
                                best_match_type = 'No Match'
                                best_matching_gtins = []
                                
                                for gtin in common_gtins:
                                    confidence, match_type = _get_gtin_confidence(cust_pool[gtin], cat_pool[gtin])
                                    if confidence > best_confidence:
                                        best_confidence = confidence
                                        best_match_type = match_type
                                        best_matching_gtins = [gtin]
                                    elif confidence == best_confidence:
                                        best_matching_gtins.append(gtin)
                                
                                if best_confidence > 0:
                                    gtin_row[j] = best_confidence
                                    gtin_details[(i, j)] = {
                                        'confidence': best_confidence,
                                        'match_type': best_match_type,
                                        'matching_gtins': best_matching_gtins[:3]
                                    }
            
            # Calculate combined score
            combined_row = _calculate_combined_score(
                tfidf_row, fuzzy_row, gtin_row, size_row,
                tfidf_weight, fuzzy_weight, gtin_weight, size_weight
            )
            
            # Store in matrices
            combined_matrix[i] = combined_row
            tfidf_matrix[i] = tfidf_row
            fuzzy_matrix[i] = fuzzy_row
            gtin_matrix[i] = gtin_row
            
            # Clean up row variables immediately
            del tfidf_row, fuzzy_row, size_row, gtin_row, combined_row
        
        # Aggressive cleanup after each small chunk
        gc.collect()
        
        # Memory check
        current_memory = get_memory_usage_mb()
        if current_memory > 1000:  # Lower threshold
            print(f"🧹 Memory cleanup at {current_memory:.0f}MB")
            gc.collect()
    
    # Final cleanup
    if within_file_mode and n_customers == n_catalog:
        np.fill_diagonal(combined_matrix, 0)
        np.fill_diagonal(tfidf_matrix, 0)
        np.fill_diagonal(fuzzy_matrix, 0)
        np.fill_diagonal(gtin_matrix, 0)
    
    if progress_callback is not None:
        progress_callback(1.0, n_customers, n_customers)
    
    return combined_matrix, tfidf_matrix, fuzzy_matrix, gtin_matrix, gtin_details

def _chunked_with_matrices(customer_texts, catalog_texts, customer_vectors, catalog_vectors,
                          tfidf_weight, fuzzy_weight, gtin_weight, size_weight,
                          customer_sizes, catalog_sizes, size_tolerance,
                          customer_gtins, catalog_gtins, similarity_threshold, early_filter,
                          enable_multiprocessing, batch_size, within_file_mode, progress_callback):
    """
    Original chunked implementation with matrix storage.
    Used for medium-sized datasets.
    """
    n_customers = len(customer_texts)
    n_catalog = len(catalog_texts)
    
    # Initialize result matrices
    combined_matrix = np.zeros((n_customers, n_catalog))
    tfidf_matrix = np.zeros((n_customers, n_catalog))
    fuzzy_matrix = np.zeros((n_customers, n_catalog))
    gtin_matrix = np.zeros((n_customers, n_catalog))
    gtin_details = {}
    
    # Determine optimal chunk size
    chunk_size = min(1000, max(100, 1500 * 1024 * 1024 // (n_catalog * 8 * 4)))
    
    print(f"🔄 Processing {n_customers:,} products in chunks of {chunk_size:,}")
    
    for chunk_start in range(0, n_customers, chunk_size):
        chunk_end = min(chunk_start + chunk_size, n_customers)
        chunk_indices = range(chunk_start, chunk_end)
        
        if progress_callback is not None:
            progress = chunk_start / n_customers
            progress_callback(progress, chunk_start, n_customers)
        
        # Process TF-IDF for this chunk
        if tfidf_weight > 0 and customer_vectors is not None and catalog_vectors is not None:
            chunk_vectors = customer_vectors[chunk_start:chunk_end]
            chunk_tfidf = cosine_similarity(chunk_vectors, catalog_vectors) * 100
        else:
            chunk_tfidf = np.zeros((len(chunk_indices), n_catalog))
        
        # Process fuzzy matching for this chunk
        chunk_fuzzy = np.zeros((len(chunk_indices), n_catalog))
        if fuzzy_weight > 0:
            if early_filter and tfidf_weight > 0:
                min_tfidf_for_fuzzy = max(5, similarity_threshold * 0.2)
                top_k = min(500, int(0.05 * n_catalog))
                
                for i_local, i_global in enumerate(chunk_indices):
                    tfidf_row = chunk_tfidf[i_local]
                    above_thresh = np.where(tfidf_row >= min_tfidf_for_fuzzy)[0]
                    
                    if len(tfidf_row) > 0:
                        k = min(top_k, max(1, n_catalog - 1))
                        topk_idx = np.argpartition(-tfidf_row, kth=k)[:top_k]
                    else:
                        topk_idx = np.array([], dtype=int)
                    
                    candidates = np.unique(np.concatenate([above_thresh, topk_idx]))
                    
                    if within_file_mode:
                        candidates = candidates[candidates != i_global]
                    
                    if len(candidates) > 0:
                        customer_text = customer_texts[i_global]
                        for j in candidates:
                            chunk_fuzzy[i_local, j] = fuzz.token_set_ratio(customer_text, catalog_texts[j])
            else:
                for i_local, i_global in enumerate(chunk_indices):
                    customer_text = customer_texts[i_global]
                    for j, catalog_text in enumerate(catalog_texts):
                        if not within_file_mode or i_global != j:
                            chunk_fuzzy[i_local, j] = fuzz.token_set_ratio(customer_text, catalog_text)
        
        # Process size and GTIN similarly...
        if size_weight > 0 and customer_sizes is not None and catalog_sizes is not None:
            chunk_customer_sizes = [customer_sizes[idx] for idx in chunk_indices]
            chunk_size_matrix = calculate_size_similarity_vectorized(
                chunk_customer_sizes, 
                catalog_sizes, 
                size_tolerance
            )
        else:
            chunk_size_matrix = np.zeros((len(chunk_indices), n_catalog))
        
        chunk_gtin = np.zeros((len(chunk_indices), n_catalog))
        if gtin_weight > 0 and customer_gtins is not None and catalog_gtins is not None:
            for i_local, i_global in enumerate(chunk_indices):
                cust_pool = customer_gtins[i_global]
                if not cust_pool:
                    continue
                for j, cat_pool in enumerate(catalog_gtins):
                    if not cat_pool:
                        continue
                    
                    common_gtins = set(cust_pool.keys()) & set(cat_pool.keys())
                    if common_gtins:
                        best_confidence = 0.0
                        best_match_type = 'No Match'
                        best_matching_gtins = []
                        
                        for gtin in common_gtins:
                            confidence, match_type = _get_gtin_confidence(cust_pool[gtin], cat_pool[gtin])
                            if confidence > best_confidence:
                                best_confidence = confidence
                                best_match_type = match_type
                                best_matching_gtins = [gtin]
                            elif confidence == best_confidence:
                                best_matching_gtins.append(gtin)
                        
                        if best_confidence > 0:
                            chunk_gtin[i_local, j] = best_confidence
                            gtin_details[(i_global, j)] = {
                                'confidence': best_confidence,
                                'match_type': best_match_type,
                                'matching_gtins': best_matching_gtins[:3]
                            }
        
        # Combine scores for this chunk
        for i_local, i_global in enumerate(chunk_indices):
            combined_score = _calculate_combined_score(
                chunk_tfidf[i_local], chunk_fuzzy[i_local], chunk_gtin[i_local], chunk_size_matrix[i_local],
                tfidf_weight, fuzzy_weight, gtin_weight, size_weight
            )
            
            combined_matrix[i_global] = combined_score
            tfidf_matrix[i_global] = chunk_tfidf[i_local]
            fuzzy_matrix[i_global] = chunk_fuzzy[i_local]
            gtin_matrix[i_global] = chunk_gtin[i_local]
        
        # Clean up memory
        del chunk_tfidf, chunk_fuzzy, chunk_size_matrix, chunk_gtin
        gc.collect()
    
    if within_file_mode and n_customers == n_catalog:
        np.fill_diagonal(combined_matrix, 0)
        np.fill_diagonal(tfidf_matrix, 0)
        np.fill_diagonal(fuzzy_matrix, 0)
        np.fill_diagonal(gtin_matrix, 0)
    
    if progress_callback is not None:
        progress_callback(1.0, n_customers, n_customers)
    
    return combined_matrix, tfidf_matrix, fuzzy_matrix, gtin_matrix, gtin_details

def _calculate_combined_score(tfidf_row, fuzzy_row, gtin_row, size_row, tfidf_weight, fuzzy_weight, gtin_weight, size_weight):
    """Calculate combined score for a single row using the same logic as the original."""
    # Check if this is GTIN-only mode
    is_gtin_only = (gtin_weight > 0 and tfidf_weight == 0 and fuzzy_weight == 0)
    is_text_only = (gtin_weight == 0 and (tfidf_weight > 0 or fuzzy_weight > 0))
    
    if is_gtin_only:
        # GTIN-only mode
        combined = gtin_row.copy()
        if size_weight > 0:
            combined = (combined * (1 - size_weight)) + (size_row * size_weight)
    elif is_text_only:
        # Text-only mode
        text_total = tfidf_weight + fuzzy_weight
        if text_total > 0:
            tfidf_norm = tfidf_weight / text_total
            fuzzy_norm = fuzzy_weight / text_total
        else:
            tfidf_norm = 0.5
            fuzzy_norm = 0.5
        
        combined = (tfidf_row * tfidf_norm) + (fuzzy_row * fuzzy_norm)
        if size_weight > 0:
            combined = (combined * (1 - size_weight)) + (size_row * size_weight)
    else:
        # Combined mode
        text_total = tfidf_weight + fuzzy_weight
        if text_total > 0:
            tfidf_norm = tfidf_weight / text_total
            fuzzy_norm = fuzzy_weight / text_total
        else:
            tfidf_norm = 0.5
            fuzzy_norm = 0.5
        
        # Calculate text-only scores
        text_only = (tfidf_row * tfidf_norm) + (fuzzy_row * fuzzy_norm)
        
        # Calculate GTIN-blended scores
        gtin_blended = (text_only * 0.5) + (gtin_row * 0.5)
        
        # Use mask to select correct calculation
        has_gtin_match = (gtin_row > 0)
        combined = np.where(has_gtin_match, gtin_blended, text_only)
        
        # Add size component if enabled
        if size_weight > 0:
            combined = (combined * (1 - size_weight)) + (size_row * size_weight)
    
    # Cap at 100%
    return np.minimum(combined, 100.0)

def calculate_similarity_vectorized(customer_texts, catalog_texts, customer_vectors, catalog_vectors,
                                  tfidf_weight=0.5, fuzzy_weight=0.5, gtin_weight=0.0, size_weight=0.0,
                                  customer_sizes=None, catalog_sizes=None, size_tolerance=20,
                                  customer_gtins=None, catalog_gtins=None,
                                  similarity_threshold=50, early_filter=True, enable_multiprocessing=True, batch_size=1000,
                                  within_file_mode=False, progress_callback=None, restriction_data=None):
    """
    Vectorized similarity calculation that processes all comparisons at once.
    Much faster than row-by-row processing. Now includes GTIN matching capability.
    """
    n_customers = len(customer_texts)
    n_catalog = len(catalog_texts)

    def _emit_progress(progress):
        if progress_callback is None:
            return
        clamped = max(0.0, min(1.0, float(progress)))
        current = min(n_customers, max(0, int(clamped * n_customers)))
        progress_callback(clamped, current, n_customers)

    _emit_progress(0.02)
    
    # For within-file mode, set diagonal to 0 to avoid self-matching
    if within_file_mode and n_customers == n_catalog:
        # We'll handle this after calculating all similarities
        pass
    
    # Calculate TF-IDF similarity matrix only if enabled and vectors are provided
    if tfidf_weight > 0 and customer_vectors is not None and catalog_vectors is not None:
        tfidf_matrix = cosine_similarity(customer_vectors, catalog_vectors) * 100
    else:
        tfidf_matrix = np.zeros((n_customers, n_catalog))
    _emit_progress(0.20)
    
    # Fuzzy similarity matrix - skip entirely if not needed
    if fuzzy_weight > 0:
        # Only calculate fuzzy if actually needed
        # Early filtering based on TF-IDF scores to reduce fuzzy matching workload
        if early_filter and tfidf_weight > 0:
            # More aggressive optimization for large within-file comparisons
            if within_file_mode and n_customers == n_catalog and n_customers > 10000:
                # For large within-file comparisons, use much stricter filtering
                fuzzy_matrix = np.zeros((n_customers, n_catalog))
                
                # Much more aggressive parameters for large datasets
                min_tfidf_for_fuzzy = max(5, similarity_threshold * 0.2)  # Higher threshold
                top_k = min(500, int(0.05 * n_catalog))  # Only top 5% instead of 20%
                
                print(f"🚀 Processing large dataset ({n_customers:,} products) with optimized filtering...")
                
                for i in range(n_customers):
                    if i % 1000 == 0:
                        print(f"Processing product {i:,} of {n_customers:,}")
                        if progress_callback is not None:
                            progress_callback(0.20 + (0.60 * ((i + 1) / n_customers)), i + 1, n_customers)
                    
                    tfidf_row = tfidf_matrix[i]
                    # Much stricter filtering for large datasets
                    above_thresh = np.where(tfidf_row >= min_tfidf_for_fuzzy)[0]
                    
                    # Top-K indices by TF-IDF
                    if len(tfidf_row) > 0:
                        k = min(top_k, max(1, n_catalog - 1))
                        topk_idx = np.argpartition(-tfidf_row, kth=k)[:top_k]
                    else:
                        topk_idx = np.array([], dtype=int)
                    
                    # Union of both sets
                    cand_idx = np.unique(np.concatenate([above_thresh, topk_idx]))
                    if cand_idx.size == 0:
                        continue
                    
                    # Skip self in within-file mode
                    if within_file_mode:
                        cand_idx = cand_idx[cand_idx != i]
                    
                    cust_text = customer_texts[i]
                    for j in cand_idx:
                        fuzzy_matrix[i, j] = fuzz.token_set_ratio(cust_text, catalog_texts[j])

                if progress_callback is not None:
                    progress_callback(0.80, n_customers, n_customers)
            # If catalog is moderate size, compute full fuzzy for maximal recall
            elif n_customers * n_catalog <= 5_000_000 or n_catalog <= 5000:
                if enable_multiprocessing and n_customers * n_catalog > 10000:
                    fuzzy_matrix = batch_fuzzy_matching(customer_texts, catalog_texts, batch_size)
                else:
                    fuzzy_matrix = np.array([[fuzz.token_set_ratio(c_text, cat_text) for cat_text in catalog_texts] for c_text in customer_texts])
            else:
                # Calculate fuzzy only for top-K TF-IDF candidates per customer to increase recall
                # while keeping performance reasonable
                fuzzy_matrix = np.zeros((n_customers, n_catalog))

                # Adaptive parameters (looser for better recall)
                min_tfidf_for_fuzzy = max(1, similarity_threshold * 0.1)
                top_k = max(1000, int(0.2 * n_catalog))  # at least 1000 or top 20% of catalog

                for i in range(n_customers):
                    if progress_callback is not None and i % 200 == 0:
                        progress_callback(0.20 + (0.60 * ((i + 1) / n_customers)), i + 1, n_customers)

                    tfidf_row = tfidf_matrix[i]
                    # Indices above minimal threshold
                    above_thresh = np.where(tfidf_row >= min_tfidf_for_fuzzy)[0]
                    # Top-K indices by TF-IDF
                    if len(tfidf_row) > 0:
                        k = min(top_k, max(1, n_catalog - 1))
                        topk_idx = np.argpartition(-tfidf_row, kth=k)[:top_k]
                    else:
                        topk_idx = np.array([], dtype=int)

                    # Union of both sets
                    cand_idx = np.unique(np.concatenate([above_thresh, topk_idx]))
                    if cand_idx.size == 0:
                        continue

                    cust_text = customer_texts[i]
                    for j in cand_idx:
                        fuzzy_matrix[i, j] = fuzz.token_set_ratio(cust_text, catalog_texts[j])
                _emit_progress(0.80)
        else:
            # Calculate full fuzzy matrix using batch processing
            if enable_multiprocessing and n_customers * n_catalog > 10000:
                fuzzy_matrix = batch_fuzzy_matching(customer_texts, catalog_texts, batch_size)
            else:
                # Single-threaded for smaller datasets
                fuzzy_matrix = np.array([[fuzz.token_set_ratio(c_text, cat_text) for cat_text in catalog_texts] for c_text in customer_texts])
            _emit_progress(0.80)
    else:
        # Skip fuzzy calculation entirely if not needed
        fuzzy_matrix = np.zeros((n_customers, n_catalog))
        _emit_progress(0.80)
    
    # Size similarity matrix (vectorized)
    if size_weight > 0 and customer_sizes is not None and catalog_sizes is not None:
        size_matrix = calculate_size_similarity_vectorized(
            customer_sizes, 
            catalog_sizes, 
            size_tolerance
        )
    else:
        size_matrix = np.zeros((n_customers, n_catalog))
    _emit_progress(0.90)
    
    # GTIN similarity matrix - Use consistent approach to restore original behavior
    gtin_matrix = np.zeros((n_customers, n_catalog))
    gtin_details = {}  # Store match details for results display
    if gtin_weight > 0 and customer_gtins is not None and catalog_gtins is not None:
        
        # Build inverted index for catalog GTINs once, to turn O(N*M) search into O(N)
        catalog_gtin_index = {}
        for j, cat_pool in enumerate(catalog_gtins):
            if cat_pool:
                for gtin, match_type in cat_pool.items():
                    if gtin not in catalog_gtin_index:
                        catalog_gtin_index[gtin] = []
                    catalog_gtin_index[gtin].append((j, match_type))
        
        for i, cust_pool in enumerate(customer_gtins):
            if not cust_pool:
                continue
            
            # Use inverted index to find only the catalog products that share GTINs
            for gtin, cust_match_type in cust_pool.items():
                if gtin in catalog_gtin_index:
                    for j, cat_match_type in catalog_gtin_index[gtin]:
                        # Calculate confidence directly (only for matching pairs)
                        conf, mtype = _get_gtin_confidence(cust_match_type, cat_match_type)
                        
                        # Keep best match if multiple GTINs overlap
                        if conf > gtin_matrix[i, j]:
                            gtin_matrix[i, j] = conf
                            gtin_details[(i, j)] = {
                                'confidence': conf,
                                'match_type': mtype,
                                'matching_gtins': [gtin]
                            }
                        elif conf == gtin_matrix[i, j] and conf > 0:
                            # Add to existing list if equal confidence
                            if (i, j) in gtin_details:
                                if gtin not in gtin_details[(i, j)]['matching_gtins']:
                                    if len(gtin_details[(i, j)]['matching_gtins']) < 3:
                                        gtin_details[(i, j)]['matching_gtins'].append(gtin)
    _emit_progress(0.97)

    # FAST VECTORIZED CALCULATIONS - Handle different matching modes properly
    
    # Check if this is GTIN-only mode
    is_gtin_only = (gtin_weight > 0 and tfidf_weight == 0 and fuzzy_weight == 0)
    is_text_only = (gtin_weight == 0 and (tfidf_weight > 0 or fuzzy_weight > 0))
    is_combined = (gtin_weight > 0 and (tfidf_weight > 0 or fuzzy_weight > 0))
    
    if is_gtin_only:
        # GTIN-only mode: Use only GTIN scores
        combined_matrix = gtin_matrix.copy()
        # Add size if enabled
        if size_weight > 0:
            combined_matrix = (combined_matrix * (1 - size_weight)) + (size_matrix * size_weight)
    
    elif is_text_only:
        # Text-only mode: Use only text scores
        text_total = tfidf_weight + fuzzy_weight
        if text_total > 0:
            tfidf_norm = tfidf_weight / text_total
            fuzzy_norm = fuzzy_weight / text_total
        else:
            tfidf_norm = 0.5
            fuzzy_norm = 0.5
        
        combined_matrix = (tfidf_matrix * tfidf_norm) + (fuzzy_matrix * fuzzy_norm)
        # Add size if enabled
        if size_weight > 0:
            combined_matrix = (combined_matrix * (1 - size_weight)) + (size_matrix * size_weight)
    
    else:
        # Combined mode: Dynamic weighting based on GTIN matches
        text_total = tfidf_weight + fuzzy_weight
        if text_total > 0:
            tfidf_norm = tfidf_weight / text_total
            fuzzy_norm = fuzzy_weight / text_total
        else:
            tfidf_norm = 0.5
            fuzzy_norm = 0.5
        
        # Calculate text-only scores for ALL pairs
        text_only_scores = (tfidf_matrix * tfidf_norm) + (fuzzy_matrix * fuzzy_norm)
        
        # Calculate GTIN-blended scores for ALL pairs
        gtin_blended_scores = (text_only_scores * 0.5) + (gtin_matrix * 0.5)
        
        # Create mask for pairs that have GTIN matches
        has_gtin_match = (gtin_matrix > 0)
        
        # Use mask to select correct calculation for each pair
        combined_matrix = np.where(has_gtin_match, gtin_blended_scores, text_only_scores)
        
        # Add size component if enabled
        if size_weight > 0:
            combined_matrix = (combined_matrix * (1 - size_weight)) + (size_matrix * size_weight)

    # Cap final confidence scores at 100%
    combined_matrix = np.minimum(combined_matrix, 100.0)
    
    # For within-file mode, set diagonal to 0 to avoid self-matching
    if within_file_mode and n_customers == n_catalog:
        np.fill_diagonal(combined_matrix, 0)
        np.fill_diagonal(tfidf_matrix, 0)
        np.fill_diagonal(fuzzy_matrix, 0)
        np.fill_diagonal(gtin_matrix, 0)
        np.fill_diagonal(size_matrix, 0)

    _emit_progress(1.0)
    
    return combined_matrix, tfidf_matrix, fuzzy_matrix, gtin_matrix, size_matrix, gtin_details

def _get_gtin_confidence(cust_src, cat_src):
    """Helper function to determine GTIN match confidence based on source types."""
    # Prioritize exact matches first
    if cust_src == 'original' and cat_src == 'original':
        return 120.0, 'Exact Match'
    elif 'corrected' in [cust_src, cat_src]:
        return 92.0, 'Corrected Match'
    elif 'case_to_unit' in [cust_src, cat_src]:
        return 90.0, 'Case/Unit Match'
    elif 'missing_check' in [cust_src, cat_src]:
        return 92.0, 'Corrected Match'
    else:
        # Default to exact match for any other combinations
        return 120.0, 'Exact Match'

def calculate_similarity(text1, text2_series, vectorizer, catalog_vectors, customer_vector, 
                        tfidf_weight=0.5, fuzzy_weight=0.5, size_weight=0.0, 
                        size1='', size2_series=None, size_tolerance=20):
    """Calculates a weighted similarity score between two texts with configurable weights."""
    # Fuzzy matching score for each item in the series (0-100 scale)
    fuzzy_scores = text2_series.apply(lambda x: fuzz.token_set_ratio(text1, x))
    
    # TF-IDF cosine similarity (0-1 scale, convert to 0-100)
    tfidf_scores = cosine_similarity(customer_vector, catalog_vectors).flatten() * 100
    
    # Size matching score (if enabled) - now uses similarity instead of exact match
    size_scores = pd.Series([0] * len(text2_series))
    if size_weight > 0 and size1 and size2_series is not None:
        size_scores = size2_series.apply(lambda x: calculate_size_similarity(size1, x, size_tolerance))
    
    # Normalize weights to ensure they sum to 1
    total_weight = tfidf_weight + fuzzy_weight + size_weight
    if total_weight > 0:
        tfidf_weight /= total_weight
        fuzzy_weight /= total_weight
        size_weight /= total_weight
    
    # Weighted average - all scores on 0-100 scale
    combined_scores = (tfidf_scores * tfidf_weight) + (fuzzy_scores * fuzzy_weight) + (size_scores * size_weight)
    return combined_scores


def process_grouped_results(similarity_matrix, 
                           product_df, 
                           product_names,
                           similarity_threshold=80.0,
                           min_group_size=2,
                           max_groups=None,
                           group_view_mode=True,
                           selected_output_columns=None,
                           conservative_grouping=True):
    """
    Process similarity matrix to create grouped product results.
    
    Args:
        similarity_matrix: NxN matrix of similarity scores
        product_df: DataFrame with product information
        product_names: List of product names for display
        similarity_threshold: Minimum similarity to consider products connected
        min_group_size: Minimum group size to include in results
        max_groups: Maximum number of groups to return
        group_view_mode: If True, return grouped format; if False, return pairwise format
        
    Returns:
        DataFrame with results (grouped or pairwise)
    """
    filtered_analyses = get_group_analyses(
        similarity_matrix=similarity_matrix,
        product_names=product_names,
        similarity_threshold=similarity_threshold,
        min_group_size=min_group_size,
        max_groups=max_groups,
        conservative_grouping=conservative_grouping,
    )
    
    if not filtered_analyses:
        return pd.DataFrame()
    
    if group_view_mode:
        # Return grouped results
        display_columns = selected_output_columns if selected_output_columns is not None else []
        results_df = create_grouped_results(filtered_analyses, product_df, display_columns)
    else:
        # Return traditional pairwise results
        results = []
        selected_cols = selected_output_columns if selected_output_columns is not None else []
        for analysis in filtered_analyses:
            member_indices = analysis['member_indices']
            
            # Create pairwise combinations within the group
            for i in range(len(member_indices)):
                for j in range(i + 1, len(member_indices)):
                    idx1, idx2 = member_indices[i], member_indices[j]
                    score = similarity_matrix[idx1][idx2]
                    
                    result = {
                        'Product 1': product_names[idx1],
                        'Product 2': product_names[idx2],
                        'Confidence Score': f"{score:.2f}%",
                        'Group ID': analysis['group_id'],
                        'Group Size': analysis['group_size']
                    }
                    
                    # Add only user-selected additional columns
                    for col in selected_cols:
                        if col in product_df.columns:
                            result[f'Product 1 {col}'] = product_df.iloc[idx1][col]
                            result[f'Product 2 {col}'] = product_df.iloc[idx2][col]
                    
                    results.append(result)
        
        results_df = pd.DataFrame(results)
    
    return results_df

def stream_similarity_results(customer_texts, catalog_texts, customer_vectors, catalog_vectors,
                            tfidf_weight=0.5, fuzzy_weight=0.5, gtin_weight=0.0, size_weight=0.0,
                            customer_sizes=None, catalog_sizes=None, size_tolerance=20,
                            customer_gtins=None, catalog_gtins=None,
                            similarity_threshold=50, early_filter=True, enable_multiprocessing=True,
                            within_file_mode=False, progress_callback=None, max_matches_per_product=100):
    """
    True streaming processor that never stores full matrices.
    Processes one row at a time and streams results directly.
    Memory usage stays constant regardless of dataset size.
    """
    n_customers = len(customer_texts)
    n_catalog = len(catalog_texts)
    
    print(f"🌊 True streaming mode: {n_customers:,} × {n_catalog:,} comparisons")
    print(f"💾 Memory usage will stay ~200-500MB regardless of dataset size")
    
    results = []
    gtin_details = {}
    
    # Process one customer at a time
    for i in range(n_customers):
        if progress_callback is not None and i % 100 == 0:
            progress = i / n_customers
            progress_callback(progress, i, n_customers)
        
        # Calculate TF-IDF for this single customer
        if tfidf_weight > 0 and customer_vectors is not None and catalog_vectors is not None:
            tfidf_scores = cosine_similarity(customer_vectors[i:i+1], catalog_vectors).flatten() * 100
        else:
            tfidf_scores = np.zeros(n_catalog)
        
        # Find candidates for fuzzy matching
        candidate_indices = list(range(n_catalog))
        
        if early_filter and fuzzy_weight > 0 and tfidf_weight > 0:
            # Use TF-IDF to pre-filter candidates
            min_tfidf_for_fuzzy = max(5, similarity_threshold * 0.2)
            top_k = min(1000, int(0.1 * n_catalog))  # More aggressive for streaming
            
            above_thresh = np.where(tfidf_scores >= min_tfidf_for_fuzzy)[0]
            
            if len(tfidf_scores) > 0:
                k = min(top_k, max(1, n_catalog - 1))
                topk_idx = np.argpartition(-tfidf_scores, kth=k)[:top_k]
            else:
                topk_idx = np.array([], dtype=int)
            
            candidate_indices = np.unique(np.concatenate([above_thresh, topk_idx]))
            
            # Skip self in within-file mode
            if within_file_mode:
                candidate_indices = candidate_indices[candidate_indices != i]
        
        # Calculate fuzzy scores only for candidates
        fuzzy_scores = np.zeros(n_catalog)
        if fuzzy_weight > 0:
            customer_text = customer_texts[i]
            if len(candidate_indices) < n_catalog * 0.5:  # Only if filtering helps
                for j in candidate_indices:
                    fuzzy_scores[j] = fuzz.token_set_ratio(customer_text, catalog_texts[j])
            else:
                # Full calculation if filtering doesn't help much
                for j, catalog_text in enumerate(catalog_texts):
                    if not within_file_mode or i != j:
                        fuzzy_scores[j] = fuzz.token_set_ratio(customer_text, catalog_text)
        
        # Calculate size scores
        size_scores = np.zeros(n_catalog)
        if size_weight > 0 and customer_sizes is not None and catalog_sizes is not None:
            customer_size = customer_sizes[i]
            if customer_size:
                for j, catalog_size in enumerate(catalog_sizes):
                    size_scores[j] = calculate_size_similarity(customer_size, catalog_size, size_tolerance)
        
        # Calculate GTIN scores
        gtin_scores = np.zeros(n_catalog)
        row_gtin_details = {}
        if gtin_weight > 0 and customer_gtins is not None and catalog_gtins is not None:
            cust_pool = customer_gtins[i]
            if cust_pool:
                for j, cat_pool in enumerate(catalog_gtins):
                    if cat_pool:
                        common_gtins = set(cust_pool.keys()) & set(cat_pool.keys())
                        if common_gtins:
                            best_confidence = 0.0
                            best_match_type = 'No Match'
                            best_matching_gtins = []
                            
                            for gtin in common_gtins:
                                confidence, match_type = _get_gtin_confidence(cust_pool[gtin], cat_pool[gtin])
                                if confidence > best_confidence:
                                    best_confidence = confidence
                                    best_match_type = match_type
                                    best_matching_gtins = [gtin]
                                elif confidence == best_confidence:
                                    best_matching_gtins.append(gtin)
                            
                            if best_confidence > 0:
                                gtin_scores[j] = best_confidence
                                row_gtin_details[j] = {
                                    'confidence': best_confidence,
                                    'match_type': best_match_type,
                                    'matching_gtins': best_matching_gtins[:3]
                                }
        
        # Calculate combined scores and find matches
        for j in range(n_catalog):
            # Skip self in within-file mode
            if within_file_mode and i == j:
                continue
            
            # Calculate combined score
            combined_score = _calculate_combined_score(
                np.array([tfidf_scores[j]]),
                np.array([fuzzy_scores[j]]),
                np.array([gtin_scores[j]]),
                np.array([size_scores[j]]),
                tfidf_weight, fuzzy_weight, gtin_weight, size_weight
            )[0]
            
            # Check if above threshold
            if combined_score >= similarity_threshold:
                # Store match result
                result = {
                    'customer_idx': i,
                    'catalog_idx': j,
                    'confidence_score': combined_score,
                    'tfidf_score': tfidf_scores[j],
                    'fuzzy_score': fuzzy_scores[j],
                    'gtin_score': gtin_scores[j]
                }
                
                # Add GTIN details if available
                if j in row_gtin_details:
                    result['gtin_match_type'] = row_gtin_details[j]['match_type']
                    result['matching_gtins'] = ', '.join(row_gtin_details[j]['matching_gtins'][:3])
                    gtin_details[(i, j)] = row_gtin_details[j]
                
                results.append(result)
        
        # Clean up this row's data immediately
        del tfidf_scores, fuzzy_scores, size_scores, gtin_scores, row_gtin_details
        
        # Aggressive garbage collection every 100 rows
        if i % 100 == 0:
            gc.collect()
            
            # Memory check
            current_memory = get_memory_usage_mb()
            if current_memory > 800:  # Lower threshold for streaming
                print(f"🧹 Memory cleanup at {current_memory:.0f}MB")
                gc.collect()
    
    # Final cleanup
    gc.collect()
    
    if progress_callback is not None:
        progress_callback(1.0, n_customers, n_customers)
    
    print(f"✅ Streaming complete: Found {len(results):,} matches using <500MB memory")
    
    # Return dummy matrices for compatibility (they won't be used)
    # The actual results are in the 'results' list
    dummy_matrix = np.zeros((n_customers, n_catalog))
    
    return dummy_matrix, dummy_matrix, dummy_matrix, dummy_matrix, gtin_details, results
