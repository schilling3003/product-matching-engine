import streamlit as st
import pandas as pd
import time
from sklearn.feature_extraction.text import TfidfVectorizer
from io import BytesIO

# Import modularized components
from src.ui import setup_sidebar, setup_column_selection
from src.processing import clean_and_standardize, calculate_similarity_memory_efficient, process_grouped_results
from src.gtin_processing import generate_gtin_quality_report
from src.threshold_explorer import compute_threshold_explorer, build_threshold_workbook, build_enhanced_threshold_workbook


def _sanitize_for_streamlit(df: pd.DataFrame) -> pd.DataFrame:
    """Return an Arrow-safe DataFrame for Streamlit display/export."""
    safe_df = df.copy()
    for col in safe_df.columns:
        if safe_df[col].dtype == "object":
            safe_df[col] = safe_df[col].apply(
                lambda x: str(x) if isinstance(x, (dict, list, set, tuple)) else x
            )
    return safe_df

def convert_streaming_results_to_dataframe(streaming_results, cleaned_customer_df, cleaned_catalog_df,
                                         column_config, is_within_file, settings, gtin_details=None):
    """Convert chunked extraction results (list of tuples) to the expected DataFrame format."""
    if not streaming_results:
        return pd.DataFrame()

    customer_display_col = column_config['customer']['product_cols'][0]
    catalog_display_col = column_config['catalog']['product_cols'][0]
    customer_out_cols = column_config['customer'].get('output_cols', [])
    catalog_out_cols = column_config['catalog'].get('output_cols', [])

    rows = []
    for rec in streaming_results:
        i, j, combined, tfidf_s, fuzzy_s, gtin_s = rec

        if is_within_file:
            entry = {
                'Product 1': cleaned_customer_df.iloc[i][customer_display_col],
                'Product 2': cleaned_catalog_df.iloc[j][catalog_display_col],
                'Confidence Score': f"{combined:.2f}%",
                'TF-IDF Score': f"{tfidf_s:.2f}%",
                'Fuzzy Score': f"{fuzzy_s:.2f}%",
            }
        else:
            entry = {
                'Customer Product': cleaned_customer_df.iloc[i][customer_display_col],
                'Catalog Product': cleaned_catalog_df.iloc[j][catalog_display_col],
                'Confidence Score': f"{combined:.2f}%",
                'TF-IDF Score': f"{tfidf_s:.2f}%",
                'Fuzzy Score': f"{fuzzy_s:.2f}%",
            }

        if gtin_s > 0:
            entry['GTIN Score'] = f"{gtin_s:.2f}%"
        if gtin_details and (i, j) in gtin_details:
            d = gtin_details[(i, j)]
            entry['GTIN Match Type'] = d['match_type']
            entry['Matching GTINs'] = ', '.join(d['matching_gtins'][:3])

        for col in customer_out_cols:
            if col in cleaned_customer_df.columns:
                entry[f'Product 1 {col}' if is_within_file else f'Customer {col}'] = cleaned_customer_df.iloc[i][col]
        for col in catalog_out_cols:
            if col in cleaned_catalog_df.columns:
                entry[f'Product 2 {col}' if is_within_file else f'Catalog {col}'] = cleaned_catalog_df.iloc[j][col]

        rows.append(entry)

    return pd.DataFrame(rows)

def main():
    """
    Main function to run the Streamlit application.
    """
    st.set_page_config(layout="wide")
    st.title("🚀 Product Matching Engine")
    st.write("Upload product lists to find matches based on a similarity score.")

    # --- UI Setup ---
    settings = setup_sidebar()
    
    # Adjust file uploaders based on matching mode
    if settings['matching_mode'] == "Find Similar Within File":
        catalog_file = st.file_uploader("📁 Product File", type=["csv", "xlsx"])
        customer_file = None  # Don't show second uploader for within-file mode
    else:
        catalog_file = st.file_uploader("📦 Your Product Catalog", type=["csv", "xlsx"])
        customer_file = st.file_uploader("👤 Customer's Product List", type=["csv", "xlsx"])

    column_config = {}
    if catalog_file and (customer_file or settings['matching_mode'] == "Find Similar Within File"):
        try:
            # Load dataframes
            catalog_df = pd.read_csv(catalog_file) if 'csv' in catalog_file.name else pd.read_excel(catalog_file)
            
            if settings['matching_mode'] == "Find Similar Within File":
                # For within-file mode, use the same dataframe for both
                customer_df = catalog_df.copy()
                # Detect and store potential restriction columns in session state
                from src.ui import smart_detect_restriction_columns
                available_restriction_columns = smart_detect_restriction_columns(catalog_df)
                st.session_state['available_restriction_columns'] = available_restriction_columns
            else:
                customer_df = pd.read_csv(customer_file) if 'csv' in customer_file.name else pd.read_excel(customer_file)
                st.session_state['available_restriction_columns'] = []
            
            # Get column configurations from UI
            column_config = setup_column_selection(catalog_df, customer_df, settings['include_size_matching'], 
                                                  settings['enable_gtin_matching'], settings['matching_mode'])

            if st.button("✨ Find Matches"):
                st.session_state.pop('match_results', None)
                with st.spinner("Processing... this may take a moment for large files."):
                    start_time = time.time()

                    # --- Data Cleaning ---
                    cleaned_catalog_df = clean_and_standardize(catalog_df, column_config['catalog'], 
                                                               settings['remove_stop_words'], settings['case_sensitive'], 
                                                               settings['include_size_in_text'])
                    cleaned_customer_df = clean_and_standardize(customer_df, column_config['customer'], 
                                                                settings['remove_stop_words'], settings['case_sensitive'], 
                                                                settings['include_size_in_text'])

                    # Mode flags used throughout processing
                    is_within_file = settings['matching_mode'] == "Find Similar Within File"

                    # --- Vectorization (only if text matching is enabled) ---
                    if settings['enable_text_matching']:
                        vectorizer = TfidfVectorizer(stop_words='english' if settings['remove_stop_words'] else None)
                        all_texts = pd.concat([cleaned_catalog_df['combined_product_name'], cleaned_customer_df['combined_product_name']]).dropna()
                        vectorizer.fit(all_texts)
                        
                        catalog_vectors = vectorizer.transform(cleaned_catalog_df['combined_product_name'].fillna(''))
                        customer_vectors = vectorizer.transform(cleaned_customer_df['combined_product_name'].fillna(''))
                    else:
                        catalog_vectors = None
                        customer_vectors = None

                    # --- Similarity Calculation ---
                    # Show status for similarity calculation on large within-file jobs
                    status_placeholder = None
                    similarity_progress_bar = None
                    similarity_status_text = None
                    progress_callback = None
                    
                    # Check if we need memory-efficient processing
                    dataset_size = len(cleaned_customer_df) * len(cleaned_catalog_df)
                    use_memory_efficient = dataset_size > 5_000_000  # 5M elements threshold
                    use_streaming = dataset_size > 50_000_000  # 50M elements for streaming
                    
                    if is_within_file and len(cleaned_customer_df) > 10000:
                        status_placeholder = st.empty()
                        status_placeholder.text("🔄 Calculating similarity matrix...")
                        similarity_progress_bar = st.progress(0, text="Calculating similarity matrix...")
                        similarity_status_text = st.empty()

                        def progress_callback(progress, current, total):
                            similarity_progress_bar.progress(
                                progress,
                                text=f"Calculating similarity matrix... {current:,} of {total:,}"
                            )
                            similarity_status_text.text(f"Processed {current:,} of {total:,} products")
                    
                    # Use memory-efficient calculation for large datasets
                    streaming_results = None
                    restriction_data = None
                    
                    # Prepare restriction data for processing
                    if is_within_file and settings.get('restrict_matches') and settings.get('selected_restrictions'):
                        restriction_data = {
                            'columns': settings['selected_restrictions'],
                            'customer_data': [],
                            'catalog_data': []
                        }
                        for col in settings['selected_restrictions']:
                            if col in customer_df.columns:
                                restriction_data['customer_data'].append(customer_df[col].fillna('').tolist())
                            else:
                                restriction_data['customer_data'].append([''] * len(customer_df))
                            if col in catalog_df.columns:
                                restriction_data['catalog_data'].append(catalog_df[col].fillna('').tolist())
                            else:
                                restriction_data['catalog_data'].append([''] * len(catalog_df))
                    
                    if use_memory_efficient:
                        result = calculate_similarity_memory_efficient(
                            customer_texts=cleaned_customer_df['combined_product_name'].fillna('').tolist(),
                            catalog_texts=cleaned_catalog_df['combined_product_name'].fillna('').tolist(),
                            customer_vectors=customer_vectors,
                            catalog_vectors=catalog_vectors,
                            tfidf_weight=settings['tfidf_weight'],
                            fuzzy_weight=settings['fuzzy_weight'],
                            gtin_weight=settings['gtin_weight'],
                            size_weight=settings['size_weight'],
                            customer_sizes=cleaned_customer_df['standardized_size'].fillna('').tolist(),
                            catalog_sizes=cleaned_catalog_df['standardized_size'].fillna('').tolist(),
                            customer_gtins=cleaned_customer_df['gtin_pool'].tolist(),
                            catalog_gtins=cleaned_catalog_df['gtin_pool'].tolist(),
                            size_tolerance=settings['size_tolerance'],
                            similarity_threshold=settings['similarity_threshold'],
                            early_filter=settings['enable_early_filtering'],
                            enable_multiprocessing=settings['enable_multiprocessing'],
                            batch_size=settings['batch_size'],
                            within_file_mode=(settings['matching_mode'] == "Find Similar Within File"),
                            progress_callback=progress_callback,
                            restriction_data=restriction_data
                        )
                        
                        # Check if streaming was used (returns 6 items instead of 5)
                        if len(result) == 6:
                            combined_matrix, tfidf_matrix, fuzzy_matrix, gtin_matrix, gtin_details, streaming_results = result
                            print(f"📊 Processing {len(streaming_results):,} streamed matches")
                        else:
                            combined_matrix, tfidf_matrix, fuzzy_matrix, gtin_matrix, gtin_details = result
                    else:
                        # Use original vectorized calculation for smaller datasets
                        from src.processing import calculate_similarity_vectorized
                        combined_matrix, tfidf_matrix, fuzzy_matrix, gtin_matrix, gtin_details = calculate_similarity_vectorized(
                            customer_texts=cleaned_customer_df['combined_product_name'].fillna('').tolist(),
                            catalog_texts=cleaned_catalog_df['combined_product_name'].fillna('').tolist(),
                            customer_vectors=customer_vectors,
                            catalog_vectors=catalog_vectors,
                            tfidf_weight=settings['tfidf_weight'],
                            fuzzy_weight=settings['fuzzy_weight'],
                            gtin_weight=settings['gtin_weight'],
                            size_weight=settings['size_weight'],
                            customer_sizes=cleaned_customer_df['standardized_size'].fillna('').tolist(),
                            catalog_sizes=cleaned_catalog_df['standardized_size'].fillna('').tolist(),
                            customer_gtins=cleaned_customer_df['gtin_pool'].tolist(),
                            catalog_gtins=cleaned_catalog_df['gtin_pool'].tolist(),
                            size_tolerance=settings['size_tolerance'],
                            similarity_threshold=settings['similarity_threshold'],
                            early_filter=settings['enable_early_filtering'],
                            enable_multiprocessing=settings['enable_multiprocessing'],
                            batch_size=settings['batch_size'],
                            within_file_mode=(settings['matching_mode'] == "Find Similar Within File"),
                            progress_callback=progress_callback,
                            restriction_data=restriction_data
                        )
                    
                    # Clear similarity calculation status
                    if status_placeholder is not None:
                        status_placeholder.empty()
                    if similarity_progress_bar is not None:
                        similarity_progress_bar.progress(1.0, text="Similarity calculation complete!")
                    if similarity_status_text is not None:
                        similarity_status_text.text("✅ Similarity calculation complete!")

                    # --- Results Processing ---
                    results = []
                    results_df = pd.DataFrame()
                    selected_group_output_cols = column_config['customer'].get('output_cols', [])
                    use_grouping = is_within_file and settings.get('group_results', False)
                    product_names = cleaned_customer_df[column_config['customer']['product_cols'][0]].tolist()
                    
                    # Check if we have streaming results
                    if streaming_results is not None:
                        print("📊 Converting streaming results to DataFrame...")
                        results_df = convert_streaming_results_to_dataframe(
                            streaming_results, cleaned_customer_df, cleaned_catalog_df,
                            column_config, is_within_file, settings, gtin_details=gtin_details
                        )
                        # Filter to top matches per product if needed
                        if settings['max_matches_per_product'] and not is_within_file:
                            # Group by customer product and take top matches
                            customer_col = 'Customer Product' if not is_within_file else 'Product 1'
                            results_df = results_df.groupby(customer_col).head(settings['max_matches_per_product'])
                    else:
                        if use_grouping:
                            # Use grouped results processing
                            product_names = cleaned_customer_df[column_config['customer']['product_cols'][0]].tolist()
                            
                            results_df = process_grouped_results(
                                similarity_matrix=combined_matrix,
                                product_df=cleaned_customer_df,
                                product_names=product_names,
                                similarity_threshold=settings['similarity_threshold'],
                                min_group_size=settings.get('min_group_size', 2),
                                max_groups=settings.get('max_groups', None),
                                group_view_mode=(settings.get('view_mode', 'Summary with Details') == 'Summary with Details'),
                                selected_output_columns=selected_group_output_cols,
                                conservative_grouping=True,
                            )
                        else:
                            # Original pairwise processing
                            # For large within-file datasets, use optimized processing
                            if is_within_file and len(cleaned_customer_df) > 10000:
                                print("🚀 Using optimized results processing for large dataset...")
                                
                                # Add progress bar for large datasets
                                progress_bar = st.progress(0, text="Processing products...")
                                status_text = st.empty()
                                
                                # Use numpy operations for faster processing
                                for i in range(len(cleaned_customer_df)):
                                    if i % 1000 == 0:
                                        progress = (i + 1) / len(cleaned_customer_df)
                                        progress_bar.progress(progress, text=f"Processing products... {i+1:,} of {len(cleaned_customer_df):,}")
                                        status_text.text(f"Processing product {i+1:,} of {len(cleaned_customer_df):,}")
                                        
                                    scores = combined_matrix[i]
                                    # Only keep indices above threshold to reduce work
                                    above_threshold = [
                                        idx for idx, value in enumerate(scores)
                                        if value >= settings['similarity_threshold'] and (not is_within_file or idx != i)
                                    ]
                                    
                                    # Get top matches from above threshold
                                    if len(above_threshold) > 0:
                                        top_n = min(settings['max_matches_per_product'], len(above_threshold))
                                        top_indices_local = sorted(
                                            above_threshold,
                                            key=lambda idx: scores[idx],
                                            reverse=True
                                        )[:top_n]
                                        
                                        for j in top_indices_local:
                                            score = scores[j]
                                            
                                            # Apply restriction filters if enabled
                                            if is_within_file and settings.get('restrict_matches') and settings.get('selected_restrictions'):
                                                skip_match = False
                                                for restriction_col in settings['selected_restrictions']:
                                                    if restriction_col in customer_df.columns and restriction_col in catalog_df.columns:
                                                        val1 = customer_df.iloc[i][restriction_col]
                                                        val2 = catalog_df.iloc[j][restriction_col]
                                                        # Handle NaN values and compare
                                                        if pd.isna(val1) or pd.isna(val2) or str(val1).lower() != str(val2).lower():
                                                            skip_match = True
                                                            break
                                                if skip_match:
                                                    continue
                                            
                                            # Use the first selected product column for display
                                            customer_display_col = column_config['customer']['product_cols'][0]
                                            catalog_display_col = column_config['catalog']['product_cols'][0]
                                            
                                            # Get the original rows for additional columns
                                            original_customer_row = customer_df.iloc[i]
                                            original_catalog_row = catalog_df.iloc[j]

                                            if is_within_file:
                                                # For within-file mode, show Product 1 and Product 2
                                                result_entry = {
                                                    'Product 1': cleaned_customer_df.iloc[i][customer_display_col],
                                                    'Product 2': cleaned_catalog_df.iloc[j][catalog_display_col],
                                                    'Confidence Score': f"{score:.2f}%",
                                                    'TF-IDF Score': f"{tfidf_matrix[i, j]:.2f}%",
                                                    'Fuzzy Score': f"{fuzzy_matrix[i, j]:.2f}%"
                                                }
                                            else:
                                                # Original two-file mode
                                                result_entry = {
                                                    'Customer Product': cleaned_customer_df.iloc[i][customer_display_col],
                                                    'Catalog Product': cleaned_catalog_df.iloc[j][catalog_display_col],
                                                    'Confidence Score': f"{score:.2f}%",
                                                    'TF-IDF Score': f"{tfidf_matrix[i, j]:.2f}%",
                                                    'Fuzzy Score': f"{fuzzy_matrix[i, j]:.2f}%"
                                                }
                                            
                                            # Skip GTIN and additional columns for speed in large datasets
                                            results.append(result_entry)
                                
                                # Complete the progress bar
                                progress_bar.progress(1.0, text="Processing complete!")
                                status_text.text("✅ Processing complete!")
                                # Clear the status text after a moment
                                time.sleep(0.5)
                                status_text.empty()
                            else:
                                # Original processing for smaller datasets
                                for i, customer_row in cleaned_customer_df.iterrows():
                                    scores = combined_matrix[i]
                                    top_indices = scores.argsort()[-settings['max_matches_per_product']:][::-1]
                                    
                                    for j in top_indices:
                                        score = scores[j]
                                        # Skip self-comparison in within-file mode
                                        if is_within_file and i == j:
                                            continue
                                        
                                        # Apply restriction filters if enabled
                                        if is_within_file and settings.get('restrict_matches') and settings.get('selected_restrictions'):
                                            skip_match = False
                                            for restriction_col in settings['selected_restrictions']:
                                                if restriction_col in customer_df.columns and restriction_col in catalog_df.columns:
                                                    val1 = customer_df.iloc[i][restriction_col]
                                                    val2 = catalog_df.iloc[j][restriction_col]
                                                    # Handle NaN values and compare
                                                    if pd.isna(val1) or pd.isna(val2) or str(val1).lower() != str(val2).lower():
                                                        skip_match = True
                                                        break
                                            if skip_match:
                                                continue
                                            
                                        if score >= settings['similarity_threshold']:
                                            # Use the first selected product column for display
                                            customer_display_col = column_config['customer']['product_cols'][0]
                                            catalog_display_col = column_config['catalog']['product_cols'][0]
                                            
                                            # Get the original rows for additional columns
                                            original_customer_row = customer_df.iloc[i]
                                            original_catalog_row = catalog_df.iloc[j]

                                            if is_within_file:
                                                # For within-file mode, show Product 1 and Product 2
                                                result_entry = {
                                                    'Product 1': customer_row[customer_display_col],
                                                    'Product 2': cleaned_catalog_df.iloc[j][catalog_display_col],
                                                    'Confidence Score': f"{score:.2f}%",
                                                    'TF-IDF Score': f"{tfidf_matrix[i, j]:.2f}%",
                                                    'Fuzzy Score': f"{fuzzy_matrix[i, j]:.2f}%"
                                                }
                                            else:
                                                # Original two-file mode
                                                result_entry = {
                                                    'Customer Product': customer_row[customer_display_col],
                                                    'Catalog Product': cleaned_catalog_df.iloc[j][catalog_display_col],
                                                    'Confidence Score': f"{score:.2f}%",
                                                    'TF-IDF Score': f"{tfidf_matrix[i, j]:.2f}%",
                                                    'Fuzzy Score': f"{fuzzy_matrix[i, j]:.2f}%"
                                                }
                                            
                                            # Add GTIN score if GTIN matching is enabled
                                            if settings['enable_gtin_matching']:
                                                gtin_score = gtin_matrix[i, j]
                                                result_entry['GTIN Score'] = f"{gtin_score:.2f}%"
                                                
                                                # Add GTIN match details if available
                                                if (i, j) in gtin_details:
                                                    details = gtin_details[(i, j)]
                                                    result_entry['GTIN Match Type'] = details['match_type']
                                                    result_entry['Matching GTINs'] = ', '.join(details['matching_gtins'][:3])  # Limit to first 3
                                            
                                            # Add additional customer columns if selected
                                            if column_config['customer']['output_cols']:
                                                for col in column_config['customer']['output_cols']:
                                                    if col in original_customer_row.index:
                                                        if is_within_file:
                                                            result_entry[f'Product 1 {col}'] = original_customer_row[col]
                                                        else:
                                                            result_entry[f'Customer {col}'] = original_customer_row[col]
                                            
                                            # Add additional catalog columns if selected
                                            if column_config['catalog']['output_cols']:
                                                for col in column_config['catalog']['output_cols']:
                                                    if col in original_catalog_row.index:
                                                        if is_within_file:
                                                            result_entry[f'Product 2 {col}'] = original_catalog_row[col]
                                                        else:
                                                            result_entry[f'Catalog {col}'] = original_catalog_row[col]
                                            
                                            results.append(result_entry)
                                
                                results_df = pd.DataFrame(results)
                    
                    end_time = time.time()
                    processing_time = end_time - start_time
                    
                    st.success(f"✅ Matching complete in {processing_time:.2f} seconds!")

                    # --- GTIN Quality Report (if enabled) ---
                    if settings['enable_gtin_matching'] and (column_config['catalog']['gtin_cols'] or column_config['customer']['gtin_cols']):
                        with st.expander("📈 GTIN Data Quality Report", expanded=False):
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                if column_config['catalog']['gtin_cols']:
                                    st.subheader("📎 Catalog GTIN Quality")
                                    catalog_report = generate_gtin_quality_report(catalog_df, column_config['catalog']['gtin_cols'])
                                    
                                    st.metric("Total Products", catalog_report['total_products'])
                                    st.metric("Products with GTINs", f"{catalog_report['products_with_gtins']} ({catalog_report['coverage_percentage']:.1f}%)")
                                    st.metric("Valid GTINs", catalog_report['valid_gtins'])
                                    st.metric("Correctable GTINs", catalog_report['correctable_gtins'])
                                    if catalog_report['invalid_gtins'] > 0:
                                        st.metric("Invalid GTINs", catalog_report['invalid_gtins'])
                            
                            with col2:
                                if column_config['customer']['gtin_cols']:
                                    st.subheader("👤 Customer GTIN Quality")
                                    customer_report = generate_gtin_quality_report(customer_df, column_config['customer']['gtin_cols'])
                                    
                                    st.metric("Total Products", customer_report['total_products'])
                                    st.metric("Products with GTINs", f"{customer_report['products_with_gtins']} ({customer_report['coverage_percentage']:.1f}%)")
                                    st.metric("Valid GTINs", customer_report['valid_gtins'])
                                    st.metric("Correctable GTINs", customer_report['correctable_gtins'])
                                    if customer_report['invalid_gtins'] > 0:
                                        st.metric("Invalid GTINs", customer_report['invalid_gtins'])
                    
                    if results_df is not None and not results_df.empty:
                        results_df = _sanitize_for_streamlit(results_df)
                        total_products = len(cleaned_customer_df)

                        threshold_summary_df = pd.DataFrame()
                        threshold_export_df = pd.DataFrame()
                        threshold_values = []
                        if use_grouping and settings.get('enable_threshold_explorer', False):
                            threshold_values, threshold_summary_df, threshold_export_df = compute_threshold_explorer(
                                similarity_matrix=combined_matrix,
                                product_names=product_names,
                                product_df=cleaned_customer_df,
                                selected_output_columns=selected_group_output_cols,
                                min_group_size=settings.get('min_group_size', 2),
                                max_groups=settings.get('max_groups', None),
                                threshold_range=settings.get('threshold_range', (50, 80)),
                                conservative_grouping=True,
                            )
                            threshold_summary_df['Singletons'] = total_products - threshold_summary_df['Products in Groups']

                        # Persist everything needed for rendering into session_state
                        st.session_state['match_results'] = {
                            'results_df': results_df,
                            'total_products': total_products,
                            'use_grouping': use_grouping,
                            'is_within_file': is_within_file,
                            'threshold_values': threshold_values,
                            'threshold_summary_df': threshold_summary_df,
                            'threshold_export_df': threshold_export_df,
                            'max_groups': settings.get('max_groups'),
                            'show_max_groups_caption': use_grouping and settings.get('max_groups') is not None,
                            'similarity_threshold': settings['similarity_threshold'],
                            # Add data for enhanced Excel export
                            'similarity_matrix': combined_matrix if use_grouping else None,
                            'product_names': product_names if use_grouping else None,
                            'cleaned_product_df': cleaned_customer_df if use_grouping else None,
                            'selected_output_columns': selected_group_output_cols if use_grouping else [],
                        }
                    else:
                        st.session_state['match_results'] = None
                        st.warning("No matches found with the current settings.")

        except Exception as e:
            st.error(f"An error occurred: {e}")

    # --- Render results from session_state (survives widget-triggered reruns) ---
    match_results = st.session_state.get('match_results')
    if match_results:
        results_df = match_results['results_df']
        total_products = match_results['total_products']
        use_grouping = match_results['use_grouping']
        is_within_file = match_results['is_within_file']
        threshold_values = match_results['threshold_values']
        threshold_summary_df = match_results['threshold_summary_df']
        threshold_export_df = match_results['threshold_export_df']
        similarity_threshold = match_results['similarity_threshold']

        if match_results.get('show_max_groups_caption'):
            st.caption(f"Showing top {match_results['max_groups']} groups by size (Maximum groups to show setting).")

        # --- Threshold Explorer ---
        current_results_df = results_df  # Start with original results
        if use_grouping and threshold_values:
            st.subheader("🎚️ Threshold Explorer")
            st.dataframe(threshold_summary_df, use_container_width=True)

            preview_threshold = st.selectbox(
                "Preview grouped results at threshold",
                threshold_values,
                index=threshold_values.index(similarity_threshold)
                if similarity_threshold in threshold_values else 0,
                key="preview_threshold",
            )
            # Update the main results_df based on selected threshold
            threshold_preview_df = threshold_export_df[
                threshold_export_df['Threshold'] == preview_threshold
            ] if not threshold_export_df.empty else pd.DataFrame()
            
            if not threshold_preview_df.empty:
                # Use the preview for display and stats
                current_results_df = threshold_preview_df

        # --- Display Stats ---
        st.subheader("📊 Match Summary")
        col1, col2, col3, col4 = st.columns(4)

        if use_grouping:
            if {'Group ID', 'Group Summary'}.issubset(current_results_df.columns):
                unique_groups = current_results_df['Group ID'].nunique()
                products_in_groups = len(current_results_df)
                avg_confidence = pd.to_numeric(current_results_df['Group Avg Similarity'], errors='coerce').mean()

                col1.metric("Total Products", f"{total_products}")
                col2.metric("Groups Found", f"{unique_groups}")
                col3.metric("Products in Groups", f"{products_in_groups}")
                col4.metric("Avg. Similarity", f"{avg_confidence:.2f}%")
            elif 'Group ID' in current_results_df.columns and 'Confidence Score' in current_results_df.columns:
                total_matches = len(current_results_df)
                avg_confidence = pd.to_numeric(current_results_df['Confidence Score'].str.replace('%', '')).mean()
                col1.metric("Total Products", f"{total_products}")
                col2.metric("Total Matches", f"{total_matches}")
                col3.metric("Groups Found", f"{current_results_df['Group ID'].nunique()}")
                col4.metric("Avg. Confidence", f"{avg_confidence:.2f}%")
        elif is_within_file:
            unique_pairs = set()
            for _, row in current_results_df.iterrows():
                pair = tuple(sorted([row['Product 1'], row['Product 2']]))
                unique_pairs.add(pair)
            products_with_matches = len(unique_pairs)
            total_matches = len(current_results_df)
            avg_confidence = pd.to_numeric(current_results_df['Confidence Score'].str.replace('%', '')).mean()

            col1.metric("Total Products", f"{total_products}")
            col2.metric("Unique Similar Pairs", f"{products_with_matches}")
            col3.metric("Total Matches", f"{total_matches}")
            col4.metric("Avg. Confidence", f"{avg_confidence:.2f}%")
        else:
            products_with_matches = current_results_df['Customer Product'].nunique()
            total_matches = len(current_results_df)
            avg_confidence = pd.to_numeric(current_results_df['Confidence Score'].str.replace('%', '')).mean()

            col1.metric("Total Customer Products", f"{total_products}")
            col2.metric("Products with Matches", f"{products_with_matches}")
            col3.metric("Total Matches Found", f"{total_matches}")
            col4.metric("Avg. Confidence", f"{avg_confidence:.2f}%")

        st.dataframe(current_results_df)

        # --- Download Buttons ---
        st.subheader("📥 Download Results")

        if use_grouping and 'Group ID' in current_results_df.columns:
            export_format = st.radio(
                "Export Format",
                ["Grouped Members", "Pairwise within Groups"],
                help="Grouped Members: One row per product with group metadata\nPairwise: Pair combinations within each group",
                key="export_format_radio",
            )

            if export_format == "Pairwise within Groups":
                export_df = process_grouped_results(
                    similarity_matrix=None,
                    product_df=None,
                    product_names=None,
                    similarity_threshold=similarity_threshold,
                    min_group_size=2,
                    max_groups=match_results['max_groups'],
                    group_view_mode=False,
                    selected_output_columns=[],
                    conservative_grouping=True,
                )
            else:
                export_df = current_results_df
        else:
            export_df = current_results_df

        export_df = _sanitize_for_streamlit(export_df)

        col1_dl, col2_dl = st.columns(2)

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Matches')
            return output.getvalue()

        with col1_dl:
            st.download_button(
                label="📄 Download as CSV",
                data=export_df.to_csv(index=False).encode('utf-8'),
                file_name='product_matches.csv',
                mime='text/csv',
            )
        with col2_dl:
            st.download_button(
                label="📊 Download as Excel",
                data=to_excel(export_df),
                file_name='product_matches.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )

        if use_grouping and not threshold_summary_df.empty and not threshold_export_df.empty:
            col1_thr, col2_thr = st.columns(2)
            
            with col1_thr:
                st.download_button(
                    label="📈 Download Threshold Explorer Workbook",
                    data=build_threshold_workbook(_sanitize_for_streamlit(threshold_export_df), threshold_summary_df),
                    file_name='group_threshold_explorer.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
            
            with col2_thr:
                # Enhanced Excel export with similarity matrix
                if match_results.get('similarity_matrix') is not None:
                    enhanced_workbook = build_enhanced_threshold_workbook(
                        similarity_matrix=match_results['similarity_matrix'],
                        product_names=match_results['product_names'],
                        product_df=match_results['cleaned_product_df'],
                        group_rows_df=_sanitize_for_streamlit(threshold_export_df),
                        summary_df=threshold_summary_df,
                        selected_output_columns=match_results.get('selected_output_columns', [])
                    )
                    st.download_button(
                        label="📊 Download Enhanced Excel for Analysis",
                        data=enhanced_workbook,
                        file_name='enhanced_threshold_analysis.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        help="Includes similarity matrix for Excel-based threshold exploration"
                    )

if __name__ == "__main__":
    main()
