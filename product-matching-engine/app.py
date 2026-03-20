import streamlit as st
import pandas as pd
import time
from sklearn.feature_extraction.text import TfidfVectorizer
from io import BytesIO

# Import modularized components
from src.ui import setup_sidebar, setup_column_selection
from src.processing import clean_and_standardize, calculate_similarity_vectorized, process_grouped_results
from src.gtin_processing import generate_gtin_quality_report

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
            else:
                customer_df = pd.read_csv(customer_file) if 'csv' in customer_file.name else pd.read_excel(customer_file)
            
            # Get column configurations from UI
            column_config = setup_column_selection(catalog_df, customer_df, settings['include_size_matching'], 
                                                  settings['enable_gtin_matching'], settings['matching_mode'])

            if st.button("✨ Find Matches"):
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
                        progress_callback=progress_callback
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
                    
                    # Check if we should use grouping (only for within-file mode)
                    use_grouping = is_within_file and settings.get('group_results', False)
                    
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
                            group_view_mode=(settings.get('view_mode', 'Summary with Details') == 'Summary with Details')
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
                        
                        # Convert results to DataFrame if not using grouping
                        if not use_grouping:
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
                        # --- Display Stats ---
                        st.subheader("📊 Match Summary")
                        total_products = len(cleaned_customer_df)
                        col1, col2, col3, col4 = st.columns(4)
                        
                        if use_grouping:
                            # For grouped results, count groups and members
                            if {'Group ID', 'Role'}.issubset(results_df.columns):
                                # Count unique groups
                                unique_groups = results_df['Group ID'].nunique()
                                # Count products in groups (exclude summary rows)
                                products_in_groups = len(results_df[results_df['Role'].isin(['Member', 'Representative'])])
                                # Calculate average confidence from numeric columns
                                if 'Confidence Score' in results_df.columns:
                                    confidence_scores = pd.to_numeric(results_df['Confidence Score'].str.replace('%', ''), errors='coerce')
                                    avg_confidence = confidence_scores.mean()
                                elif 'Avg Similarity' in results_df.columns:
                                    # Use Avg Similarity from summary rows
                                    avg_sim = results_df[results_df['Role'] == 'Representative']['Avg Similarity']
                                    avg_confidence = pd.to_numeric(avg_sim.str.replace('%', ''), errors='coerce').mean()
                                else:
                                    avg_confidence = 0
                                
                                col1.metric("Total Products", f"{total_products}")
                                col2.metric("Groups Found", f"{unique_groups}")
                                col3.metric("Products in Groups", f"{products_in_groups}")
                                col4.metric("Avg. Similarity", f"{avg_confidence:.2f}%")
                            elif 'Group ID' in results_df.columns and 'Confidence Score' in results_df.columns:
                                # Pairwise within groups
                                total_matches = len(results_df)
                                avg_confidence = pd.to_numeric(results_df['Confidence Score'].str.replace('%', '')).mean()
                                col1.metric("Total Products", f"{total_products}")
                                col2.metric("Total Matches", f"{total_matches}")
                                col3.metric("Groups Found", f"{results_df['Group ID'].nunique()}")
                                col4.metric("Avg. Confidence", f"{avg_confidence:.2f}%")
                        elif is_within_file:
                            # For within-file mode, count unique product pairs
                            unique_pairs = set()
                            for _, row in results_df.iterrows():
                                pair = tuple(sorted([row['Product 1'], row['Product 2']]))
                                unique_pairs.add(pair)
                            products_with_matches = len(unique_pairs)
                            total_matches = len(results_df)
                            avg_confidence = pd.to_numeric(results_df['Confidence Score'].str.replace('%', '')).mean()
                            
                            col1.metric("Total Products", f"{total_products}")
                            col2.metric("Unique Similar Pairs", f"{products_with_matches}")
                            col3.metric("Total Matches", f"{total_matches}")
                            col4.metric("Avg. Confidence", f"{avg_confidence:.2f}%")
                        else:
                            # Original two-file mode
                            products_with_matches = results_df['Customer Product'].nunique()
                            total_matches = len(results_df)
                            avg_confidence = pd.to_numeric(results_df['Confidence Score'].str.replace('%', '')).mean()
                            
                            col1.metric("Total Customer Products", f"{total_products}")
                            col2.metric("Products with Matches", f"{products_with_matches}")
                            col3.metric("Total Matches Found", f"{total_matches}")
                            col4.metric("Avg. Confidence", f"{avg_confidence:.2f}%")

                        st.dataframe(results_df)

                        # --- Download Buttons ---
                        st.subheader("📥 Download Results")
                        
                        # Add export format options for grouped results
                        if use_grouping and 'Group ID' in results_df.columns:
                            export_format = st.radio(
                                "Export Format",
                                ["Grouped Format", "Flat Format with Group IDs"],
                                help="Grouped Format: Shows groups with representative and members\nFlat Format: One row per product with group IDs"
                            )
                            
                            if export_format == "Flat Format with Group IDs":
                                from src.product_grouping import export_groups_flat
                                # Get analyses for flat export
                                product_names = cleaned_customer_df[column_config['customer']['product_cols'][0]].tolist()
                                from src.product_grouping import find_product_groups, analyze_groups
                                groups = find_product_groups(combined_matrix, threshold=settings['similarity_threshold'])
                                analyses = analyze_groups(combined_matrix, groups, product_names)
                                filtered_analyses = [a for a in analyses if a['group_size'] >= settings.get('min_group_size', 2)]
                                export_df = export_groups_flat(filtered_analyses, cleaned_customer_df, column_config['customer']['output_cols'])
                            else:
                                export_df = results_df
                        else:
                            export_df = results_df
                        
                        col1_dl, col2_dl = st.columns(2)
                        
                        @st.cache_data
                        def to_excel(df):
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                df.to_excel(writer, index=False, sheet_name='Matches')
                            processed_data = output.getvalue()
                            return processed_data

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
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                    else:
                        st.warning("No matches found with the current settings.")

        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
