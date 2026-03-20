import re
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from thefuzz import fuzz
from concurrent.futures import ProcessPoolExecutor
import multiprocessing as mp

from .config import STOP_WORDS, UNIT_CONVERSION_MAP
from .gtin_processing import consolidate_gtin_columns, calculate_gtin_match_confidence
from .product_grouping import find_product_groups, analyze_groups, create_grouped_results, export_groups_flat, filter_groups

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
    Optimized batch fuzzy matching using multiprocessing.
    Returns a matrix of fuzzy scores.
    """
    def fuzzy_chunk(args):
        customer_chunk, catalog_texts, start_idx = args
        results = []
        for i, customer_text in enumerate(customer_chunk):
            scores = [fuzz.token_set_ratio(customer_text, catalog_text) for catalog_text in catalog_texts]
            results.append((start_idx + i, scores))
        return results
    
    # Split customer texts into chunks for parallel processing
    chunks = []
    for i in range(0, len(customer_texts), chunk_size):
        chunk = customer_texts[i:i + chunk_size]
        chunks.append((chunk, catalog_texts, i))
    
    # Use multiprocessing for CPU-intensive fuzzy matching
    num_cores = min(mp.cpu_count(), len(chunks))
    if num_cores > 1 and len(customer_texts) > 100:  # Only use multiprocessing for larger datasets
        with ProcessPoolExecutor(max_workers=num_cores) as executor:
            chunk_results = list(executor.map(fuzzy_chunk, chunks))
        
        # Combine results
        fuzzy_matrix = np.zeros((len(customer_texts), len(catalog_texts)))
        for chunk_result in chunk_results:
            for customer_idx, scores in chunk_result:
                fuzzy_matrix[customer_idx] = scores
    else:
        # Single-threaded for smaller datasets to avoid overhead
        fuzzy_matrix = np.zeros((len(customer_texts), len(catalog_texts)))
        for i, customer_text in enumerate(customer_texts):
            for j, catalog_text in enumerate(catalog_texts):
                fuzzy_matrix[i, j] = fuzz.token_set_ratio(customer_text, catalog_text)
    
    return fuzzy_matrix

def calculate_similarity_vectorized(customer_texts, catalog_texts, customer_vectors, catalog_vectors,
                                  tfidf_weight=0.5, fuzzy_weight=0.5, gtin_weight=0.0, size_weight=0.0,
                                  customer_sizes=None, catalog_sizes=None, size_tolerance=20,
                                  customer_gtins=None, catalog_gtins=None,
                                  similarity_threshold=50, early_filter=True, enable_multiprocessing=True, batch_size=1000,
                                  within_file_mode=False, progress_callback=None):
    """
    Vectorized similarity calculation that processes all comparisons at once.
    Much faster than row-by-row processing. Now includes GTIN matching capability.
    """
    n_customers = len(customer_texts)
    n_catalog = len(catalog_texts)
    
    # For within-file mode, set diagonal to 0 to avoid self-matching
    if within_file_mode and n_customers == n_catalog:
        # We'll handle this after calculating all similarities
        pass
    
    # Calculate TF-IDF similarity matrix only if enabled and vectors are provided
    if tfidf_weight > 0 and customer_vectors is not None and catalog_vectors is not None:
        tfidf_matrix = cosine_similarity(customer_vectors, catalog_vectors) * 100
    else:
        tfidf_matrix = np.zeros((n_customers, n_catalog))
    
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
                            progress_callback((i + 1) / n_customers, i + 1, n_customers)
                    
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
                    progress_callback(1.0, n_customers, n_customers)
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
        else:
            # Calculate full fuzzy matrix using batch processing
            if enable_multiprocessing and n_customers * n_catalog > 10000:
                fuzzy_matrix = batch_fuzzy_matching(customer_texts, catalog_texts, batch_size)
            else:
                # Single-threaded for smaller datasets
                fuzzy_matrix = np.array([[fuzz.token_set_ratio(c_text, cat_text) for cat_text in catalog_texts] for c_text in customer_texts])
    else:
        # Skip fuzzy calculation entirely if not needed
        fuzzy_matrix = np.zeros((n_customers, n_catalog))
    
    # Size similarity matrix (vectorized)
    size_matrix = np.zeros((n_customers, n_catalog))
    if size_weight > 0 and customer_sizes is not None and catalog_sizes is not None:
        for i, customer_size in enumerate(customer_sizes):
            if customer_size:
                for j, catalog_size in enumerate(catalog_sizes):
                    size_matrix[i, j] = calculate_size_similarity(customer_size, catalog_size, size_tolerance)
    
    # GTIN similarity matrix - Use consistent approach to restore original behavior
    gtin_matrix = np.zeros((n_customers, n_catalog))
    gtin_details = {}  # Store match details for results display
    if gtin_weight > 0 and customer_gtins is not None and catalog_gtins is not None:
        
        # Use simple, reliable nested loop approach (same logic as original)
        for i, cust_pool in enumerate(customer_gtins):
            if not cust_pool:
                continue
            for j, cat_pool in enumerate(catalog_gtins):
                if not cat_pool:
                    continue
                
                # Find intersection of GTIN keys (same as original set intersection)
                common_gtins = set(cust_pool.keys()) & set(cat_pool.keys())
                if common_gtins:
                    # Find the best confidence match among all common GTINs
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
                        gtin_matrix[i, j] = best_confidence
                        gtin_details[(i, j)] = {
                            'confidence': best_confidence,
                            'match_type': best_match_type,
                            'matching_gtins': best_matching_gtins[:3]  # Limit to first 3
                        }

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
    
    return combined_matrix, tfidf_matrix, fuzzy_matrix, gtin_matrix, gtin_details

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
                           group_view_mode=True):
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
    # Find product groups using connected components
    groups = find_product_groups(similarity_matrix, threshold=similarity_threshold)
    
    if not groups:
        return pd.DataFrame()
    
    # Analyze groups to get statistics
    analyses = analyze_groups(similarity_matrix, groups, product_names)
    
    # Filter groups based on criteria
    filtered_analyses = filter_groups(
        analyses, 
        min_group_size=min_group_size,
        max_groups=max_groups,
        sort_by='size'
    )
    
    if not filtered_analyses:
        return pd.DataFrame()
    
    if group_view_mode:
        # Return grouped results
        display_columns = [
            col for col in product_df.columns
            if col not in ['combined_product_name', 'gtin_pool']
        ]
        results_df = create_grouped_results(filtered_analyses, product_df, display_columns)
    else:
        # Return traditional pairwise results
        results = []
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
                    
                    # Add additional columns from product_df
                    for col in product_df.columns:
                        if col not in ['combined_product_name']:
                            result[f'Product 1 {col}'] = product_df.iloc[idx1][col]
                            result[f'Product 2 {col}'] = product_df.iloc[idx2][col]
                    
                    results.append(result)
        
        results_df = pd.DataFrame(results)
    
    return results_df
