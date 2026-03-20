import streamlit as st
import pandas as pd
import numpy as np
import time
from sklearn.feature_extraction.text import TfidfVectorizer
from io import BytesIO

# Import modularized components
from src.ui import setup_sidebar, setup_column_selection
from src.processing import clean_and_standardize, calculate_similarity_vectorized
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
        catalog_file = st.file_uploader("� Product File", type=["csv", "xlsx"])
        customer_file = None  # Hide second file uploader
    else:
        catalog_file = st.file_uploader("� Your Product Catalog", type=["csv", "xlsx"])
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
                        within_file_mode=(settings['matching_mode'] == "Find Similar Within File")
                    )

                    # --- Results Processing ---
                    results = []
                    is_within_file = settings['matching_mode'] == "Find Similar Within File"
                    
                    # For large within-file datasets, use optimized processing
                    if is_within_file and len(cleaned_customer_df) > 10000:
                        print("🚀 Using optimized results processing for large dataset...")
                        
                        # Use numpy operations for faster processing
                        for i in range(len(cleaned_customer_df)):
                            if i % 1000 == 0:
                                print(f"Processing results for product {i:,} of {len(cleaned_customer_df):,}")
                                
                            scores = combined_matrix[i]
                            # Only get indices above threshold to reduce work
                            above_threshold = np.where(scores >= settings['similarity_threshold'])[0]
                            
                            # Skip self in within-file mode
                            if is_within_file:
                                above_threshold = above_threshold[above_threshold != i]
                            
                            # Get top matches from above threshold
                            if len(above_threshold) > 0:
                                # Get scores for these indices
                                threshold_scores = scores[above_threshold]
                                # Get top indices within threshold
                                top_n = min(settings['max_matches_per_product'], len(threshold_scores))
                                top_indices_local = above_threshold[np.argsort(threshold_scores)[-top_n:][::-1]]
                                
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
                    
                    if results:
                        results_df = pd.DataFrame(results)

                        # --- Display Stats ---
                        st.subheader("📊 Match Summary")
                        total_products = len(cleaned_customer_df)
                        
                        if is_within_file:
                            # For within-file mode, count unique product pairs
                            unique_pairs = set()
                            for result in results:
                                pair = tuple(sorted([result['Product 1'], result['Product 2']]))
                                unique_pairs.add(pair)
                            products_with_matches = len(unique_pairs)
                        else:
                            # Original two-file mode
                            products_with_matches = results_df['Customer Product'].nunique()
                            
                        total_matches = len(results_df)
                        # Convert confidence score to numeric for calculation
                        avg_confidence = pd.to_numeric(results_df['Confidence Score'].str.replace('%', '')).mean()

                        col1, col2, col3, col4 = st.columns(4)
                        if is_within_file:
                            col1.metric("Total Products", f"{total_products}")
                            col2.metric("Unique Similar Pairs", f"{products_with_matches}")
                            col3.metric("Total Matches", f"{total_matches}")
                            col4.metric("Avg. Confidence", f"{avg_confidence:.2f}%")
                        else:
                            col1.metric("Total Customer Products", f"{total_customer_products}")
                            col2.metric("Products with Matches", f"{products_with_matches}")
                            col3.metric("Total Matches Found", f"{total_matches}")
                            col4.metric("Avg. Confidence", f"{avg_confidence:.2f}%")

                        st.dataframe(results_df)

                        # --- Download Buttons ---
                        st.subheader("📥 Download Results")
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
                                data=results_df.to_csv(index=False).encode('utf-8'),
                                file_name='product_matches.csv',
                                mime='text/csv',
                            )
                        with col2_dl:
                            st.download_button(
                                label="📊 Download as Excel",
                                data=to_excel(results_df),
                                file_name='product_matches.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                    else:
                        st.warning("No matches found with the current settings.")

        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
