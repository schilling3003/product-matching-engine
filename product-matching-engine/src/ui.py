import streamlit as st
import pandas as pd
from .gtin_processing import smart_detect_gtin_columns

def smart_detect_product_name_columns(df):
    """Smartly detects potential product name/description columns."""
    potential_names = ['product_name', 'description', 'short_name', 'long_name', 'name', 'title', 'product', 'print name', 'item name']
    return [col for col in df.columns if col.lower() in potential_names]

def setup_sidebar():
    """Sets up the Streamlit sidebar with all the matching settings."""
    with st.sidebar:
        st.header("⚙️ Matching Settings")
        
        # === MATCHING MODE SELECTION ===
        st.subheader("🔄 Matching Mode")
        
        matching_mode = st.radio(
            "Choose matching mode:",
            ["Match Between Files", "Find Similar Within File"],
            help="Match Between Files: Compare two different product lists\nFind Similar Within File: Find similar products in the same file"
        )
        
        # === GROUPING OPTIONS (FOR WITHIN FILE MODE) ===
        if matching_mode == "Find Similar Within File":
            st.subheader("🔗 Grouping Options")
            
            group_results = st.checkbox(
                "Group similar products",
                value=True,
                help="Group all similar products together instead of showing pairwise matches. This shows the complete picture of product duplicates."
            )
            
            if group_results:
                min_group_size = st.slider(
                    "Minimum group size",
                    min_value=2, max_value=10, value=2,
                    help="Minimum number of products needed to form a group."
                )
                
                max_groups = st.selectbox(
                    "Maximum groups to show",
                    [10, 25, 50, 100, "All"],
                    index=2,
                    help="Limits the number of groups returned (ranked by group size). This does not cap members per group."
                )
                
                if max_groups == "All":
                    max_groups = None
                else:
                    max_groups = int(max_groups)
                
                view_mode = st.radio(
                    "Group view mode",
                    ["Summary with Details", "Pairwise within Groups"],
                    help="Summary with Details: Shows representative products with expandable details\nPairwise within Groups: Shows all pairwise matches within each group"
                )

                enable_threshold_explorer = st.checkbox(
                    "Enable threshold explorer",
                    value=False,
                    help="Analyze how grouping changes across similarity thresholds without rerunning full preprocessing."
                )
                threshold_range = st.slider(
                    "Threshold explorer range",
                    min_value=40,
                    max_value=95,
                    value=(50, 80),
                    step=5,
                    disabled=not enable_threshold_explorer,
                )
            else:
                min_group_size = 2
                max_groups = None
                view_mode = "Pairwise within Groups"
                enable_threshold_explorer = False
                threshold_range = (50, 80)
        else:
            group_results = False
            min_group_size = 2
            max_groups = None
            view_mode = "Pairwise within Groups"
            enable_threshold_explorer = False
            threshold_range = (50, 80)
        
        # === SIMPLE SETTINGS FOR EVERYONE ===
        st.subheader("🎯 Basic Settings")
        
        # Matching Strictness (optimized for food items)
        strictness = st.select_slider(
            "How strict should matching be?",
            options=["Very Lenient", "Lenient", "Balanced", "Strict", "Very Strict"],
            value="Lenient",
            help="Food products often have variations in naming. 'Lenient' is recommended for specialty foods."
        )
        
        # Convert strictness to threshold (much more lenient for food items)
        strictness_map = {
            "Very Lenient": 40,  # Catch very loose matches
            "Lenient": 50,       # Good for food variations
            "Balanced": 60,      # Still quite permissive
            "Strict": 70,        # More selective
            "Very Strict": 80    # Only close matches
        }
        similarity_threshold = strictness_map[strictness]
        
        # Matching method toggles
        st.subheader("🔍 Matching Methods")
        
        enable_text_matching = st.checkbox(
            "Enable Text Matching",
            value=True,
            key="text_matching_basic",
            help="Use product names and descriptions for matching."
        )
        
        enable_gtin_matching = st.checkbox(
            "Enable GTIN Matching",
            value=True,
            key="gtin_matching_basic",
            help="Use GTIN/UPC/Barcode data for matching. Handles messy data with missing check digits and mixed formats."
        )
        
        # Matching Focus - Simple and Direct (only show if text matching is enabled)
        if enable_text_matching:
            matching_focus = st.select_slider(
                "What's more important for text matching?",
                options=["Exact Spelling", "Mostly Spelling", "Balanced", "Mostly Meaning", "Product Meaning"],
                value="Balanced",
                help="Choose whether to focus on exact text matching or understanding product meanings."
            )
            
            # Convert focus to weights - these will be adjusted based on GTIN presence
            focus_weights = {
                "Exact Spelling": (0.1, 0.9),      # 10% meaning, 90% spelling
                "Mostly Spelling": (0.3, 0.7),     # 30% meaning, 70% spelling  
                "Balanced": (0.5, 0.5),             # 50% meaning, 50% spelling
                "Mostly Meaning": (0.7, 0.3),      # 70% meaning, 30% spelling
                "Product Meaning": (0.9, 0.1)      # 90% meaning, 10% spelling
            }
            base_tfidf_weight, base_fuzzy_weight = focus_weights[matching_focus]
        else:
            base_tfidf_weight, base_fuzzy_weight = 0.0, 0.0
            matching_focus = "Balanced"  # Default for display
        
        # --- Correct, Sequential Weight Calculation ---

        # 1. Determine the base weights for the main methods (Text vs. GTIN)
        if enable_text_matching and enable_gtin_matching:
            text_weight_total = 0.5
            gtin_weight = 0.5
        elif enable_text_matching:
            text_weight_total = 1.0
            gtin_weight = 0.0
        elif enable_gtin_matching:
            text_weight_total = 0.0
            gtin_weight = 1.0
        else:
            text_weight_total = 0.0
            gtin_weight = 0.0

        # 2. Split the 'text_weight' portion between TF-IDF and Fuzzy
        tfidf_weight = base_tfidf_weight * text_weight_total
        fuzzy_weight = base_fuzzy_weight * text_weight_total

        # Size matching toggle
        include_size_matching = st.checkbox(
            "Include size/weight in matching",
            value=False,
            key="size_matching_basic",
            help="When enabled, products with similar sizes get bonus points."
        )
        
        if include_size_matching:
            size_importance = st.select_slider(
                "How important is size matching?",
                options=["Minor Factor", "Moderate Factor", "Major Factor"],
                value="Moderate Factor",
                help="Choose how much weight to give to size similarity."
            )
            
            size_tolerance = st.slider(
                "Size tolerance (%)",
                min_value=5, max_value=50, value=20, step=5,
                help="How close sizes need to be. 20% means 100ml matches 80-120ml."
            )
            
            # Convert size importance to weight
            size_weights = {
                "Minor Factor": 0.1,
                "Moderate Factor": 0.2, 
                "Major Factor": 0.3
            }
            size_weight = size_weights[size_importance]
            
            # The processing module will handle the final normalization with size
        else:
            size_tolerance = 20
            size_weight = 0.0
            size_importance = "Moderate Factor"  # For display purposes
        
        # Results limit (simplified) - Only show for pairwise mode
        if not (matching_mode == "Find Similar Within File" and group_results):
            max_matches_per_product = st.selectbox(
                "How many matches to show per product?",
                [1, 3, 5, 10, 20],
                index=2,  # Default to 5
                help="Limit the number of matches shown for each customer product."
            )
        else:
            max_matches_per_product = 5  # Default value for grouped mode
        
        # Simple explanation of current settings
        threshold_desc = f"matches need {similarity_threshold}%+ similarity"
        
        # Build method description
        methods = []
        if enable_text_matching:
            methods.append(f"text ({matching_focus.lower()})")
        if enable_gtin_matching:
            methods.append("GTIN")
        if include_size_matching:
            methods.append(f"size ({size_importance.lower()})")
        
        methods_desc = " + ".join(methods) if methods else "no methods selected"
        
        if matching_mode == "Find Similar Within File" and group_results:
            result_scope_desc = f"up to {max_groups if max_groups is not None else 'all'} groups"
        else:
            result_scope_desc = f"up to {max_matches_per_product} matches per product"

        st.info(f"🍽️ **Setup:** {strictness.lower()} ({threshold_desc}) • {methods_desc} • {result_scope_desc}")
        
        # Method-specific tips
        if enable_text_matching and enable_gtin_matching:
            st.success("✅ Using both text and GTIN matching provides the most accurate results.")
        elif enable_text_matching:
            if matching_focus in ["Exact Spelling", "Mostly Spelling"]:
                st.info("🎯 Text-only matching focusing on exact spelling - good for finding products with similar names.")
            elif matching_focus in ["Mostly Meaning", "Product Meaning"]:
                st.info("💡 Text-only matching focusing on product meaning - good for finding similar items with different names.")
            else:
                st.info("📝 Text-only matching with balanced approach.")
        elif enable_gtin_matching:
            st.info("🔢 GTIN-only matching - most accurate when barcode data is available and clean.")
        else:
            st.warning("⚠️ No matching methods selected - please enable at least one method.")
        
        # General warnings
        if similarity_threshold > 70:
            st.warning("⚠️ High strictness may miss valid food variations (e.g., 'Olive Oil' vs 'Extra Virgin Olive Oil')")
        elif similarity_threshold < 45:
            st.warning("⚠️ Very low strictness may include unrelated food items")
        
        if not enable_text_matching and not enable_gtin_matching:
            st.error("❌ At least one matching method must be enabled!")
        
        # === ADVANCED SETTINGS (COLLAPSIBLE) ===
        with st.expander("🔧 Advanced Settings (Optional)", expanded=False):
            st.markdown("*For users who want fine-grained control*")
            
            # Override basic settings if user wants manual control
            manual_control = st.checkbox(
                "Enable Manual Control", 
                value=False,
                key="manual_control_toggle",
                help="Override the simple settings above with manual controls."
            )
            
            if manual_control:
                st.warning("⚠️ Manual mode enabled - basic settings above are ignored.")
                
                # Manual threshold
                similarity_threshold = st.number_input(
                    "Similarity Threshold (%)", 
                    min_value=0, max_value=100, value=similarity_threshold, step=1,
                    help="Minimum similarity score required for a match."
                )
                
                # Manual method toggles
                st.markdown("**Matching Methods:**")
                enable_text_matching = st.checkbox(
                    "Enable Text Matching", 
                    value=enable_text_matching,
                    key="text_matching_manual",
                    help="Use product names and descriptions for matching."
                )
                
                enable_gtin_matching = st.checkbox(
                    "Enable GTIN Matching", 
                    value=enable_gtin_matching,
                    key="gtin_matching_manual",
                    help="Use GTIN/UPC/Barcode data for matching."
                )
                
                # Manual algorithm weights
                st.markdown("**Algorithm Weights:**")
                tfidf_weight = st.slider(
                    "Semantic Matching (TF-IDF)", 
                    min_value=0.0, max_value=1.0, value=tfidf_weight, step=0.1,
                    help="How much to focus on word meanings and importance."
                )
                
                fuzzy_weight = st.slider(
                    "Text Similarity (Fuzzy)", 
                    min_value=0.0, max_value=1.0, value=fuzzy_weight, step=0.1,
                    help="How much to focus on exact text matching and typos."
                )
                
                gtin_weight = st.slider(
                    "GTIN Matching", 
                    min_value=0.0, max_value=1.0, value=gtin_weight, step=0.1,
                    help="How much to focus on GTIN/UPC/Barcode matching."
                )
                
                size_weight = st.slider(
                    "Size Matching", 
                    min_value=0.0, max_value=1.0, value=size_weight, step=0.1,
                    help="How much to consider size similarity."
                )
                
                if size_weight > 0:
                    size_tolerance = st.slider(
                        "Size Tolerance (%)", 
                        min_value=1, max_value=50, value=size_tolerance, step=1,
                        help="How close sizes need to be to match."
                    )
                
                # Manual results control
                max_matches_per_product = st.number_input(
                    "Max Matches per Product", 
                    min_value=1, max_value=50, value=max_matches_per_product, step=1
                )
            
            # Text processing options
            st.markdown("**Text Processing:**")
            remove_stop_words = st.checkbox(
                "Remove common words (the, and, of, etc.)", 
                value=True,
                key="remove_stop_words_manual",
                help="Usually improves matching by ignoring unimportant words."
            )
            
            case_sensitive = st.checkbox(
                "Case sensitive matching", 
                value=False,
                key="case_sensitive_manual",
                help="Treat 'Apple' and 'apple' as different (usually not recommended)."
            )
            
            # Advanced filtering
            st.markdown("**Advanced Filtering:**")
            min_tfidf_score = st.slider(
                "Minimum Semantic Score", 
                min_value=0, max_value=100, value=0, step=5,
                help="Filter out matches with poor semantic similarity."
            )
            
            min_fuzzy_score = st.slider(
                "Minimum Text Score", 
                min_value=0, max_value=100, value=0, step=5,
                help="Filter out matches with poor text similarity."
            )
            
            # Results sorting
            sort_by = st.selectbox(
                "Sort Results By",
                ["Confidence Score (Descending)", "Confidence Score (Ascending)", "Customer Product Name", "Catalog Product Name"]
            )
            
            enable_description_boost = st.checkbox(
                "Description matching boost", 
                value=True,
                key="description_boost_manual",
                help="Give extra weight when descriptions also match well."
            )
            
            # Performance optimization settings
            st.markdown("**Performance Optimization:**")
            enable_multiprocessing = st.checkbox(
                "Enable multiprocessing for large datasets", 
                value=True,
                key="multiprocessing_manual",
                help="Use multiple CPU cores for faster processing (recommended for 1000+ products)."
            )
            
            enable_early_filtering = st.checkbox(
                "Enable smart filtering", 
                value=True,
                key="early_filtering_manual",
                help="Skip obviously poor matches to speed up processing (recommended)."
            )
            
            batch_size = st.slider(
                "Processing batch size", 
                min_value=100, max_value=5000, value=1000, step=100,
                help="Larger batches use more memory but may be faster for very large datasets."
            )
        
        # Set defaults for advanced options if not in manual mode
        if 'manual_control' not in locals() or not manual_control:
            remove_stop_words = True
            case_sensitive = False
            min_tfidf_score = 0
            min_fuzzy_score = 0
            sort_by = "Confidence Score (Descending)"
            enable_description_boost = True
            enable_multiprocessing = True
            enable_early_filtering = True
            batch_size = 1000
            include_size_in_text = False  # Always disable size in text matching

    return {
        'matching_mode': matching_mode,
        'group_results': group_results,
        'min_group_size': min_group_size,
        'max_groups': max_groups,
        'view_mode': view_mode,
        'enable_threshold_explorer': enable_threshold_explorer,
        'threshold_range': threshold_range,
        'similarity_threshold': similarity_threshold,
        'enable_text_matching': enable_text_matching,
        'enable_gtin_matching': enable_gtin_matching,
        'tfidf_weight': tfidf_weight,
        'fuzzy_weight': fuzzy_weight,
        'gtin_weight': gtin_weight,
        'size_weight': size_weight,
        'size_tolerance': size_tolerance,
        'max_matches_per_product': max_matches_per_product,
        'remove_stop_words': remove_stop_words,
        'case_sensitive': case_sensitive,
        'min_tfidf_score': min_tfidf_score,
        'min_fuzzy_score': min_fuzzy_score,
        'sort_by': sort_by,
        'enable_description_boost': enable_description_boost,
        'enable_multiprocessing': enable_multiprocessing,
        'enable_early_filtering': enable_early_filtering,
        'batch_size': batch_size,
        'include_size_in_text': include_size_in_text,
        'include_size_matching': include_size_matching
    }

def setup_column_selection(catalog_df, customer_df=None, include_size_matching=False, enable_gtin_matching=True, matching_mode="Match Between Files"):
    """Sets up the column selection UI for catalog and customer files."""
    st.header("📋 Column Configuration")
    
    if matching_mode == "Find Similar Within File":
        st.info("🔍 **Within File Mode**: Configuring columns for finding similar products within the same file")
        
        catalog_cols = list(catalog_df.columns)
        
        # Product Name Columns (Multi-select)
        st.write("**Product Name / Description**")
        suggested_product_cols = smart_detect_product_name_columns(catalog_df)
        if suggested_product_cols:
            st.info(f"💡 Detected potential name/description columns: {', '.join(suggested_product_cols)}")
        
        catalog_product_cols = st.multiselect(
            "Product Name / Description Columns (Required)",
            catalog_cols,
            default=suggested_product_cols,
            key="catalog_product_cols",
            help="Select all columns containing product names or descriptions. They will be combined for matching."
        )
        
        if not catalog_product_cols:
            st.warning("⚠️ Please select at least one product name/description column.")
        
        # Size Configuration - Only show if size matching is enabled
        catalog_size_col = None
        catalog_size_value_col = None
        catalog_size_unit_col = None
        
        if include_size_matching:  # Only show size configuration if size matching is enabled
            st.write("**Size Configuration**")
            catalog_size_type = st.radio(
                "How is size data stored?",
                ["No size data", "Combined size column", "Separate value and unit columns"],
                key="catalog_size_type"
            )
            
            if catalog_size_type == "Combined size column":
                size_col_options = ["None"] + catalog_cols
                catalog_size_col = st.selectbox(
                    "Size Column", 
                    size_col_options,
                    index=size_col_options.index('size') if 'size' in catalog_cols else 0,
                    key="catalog_size"
                )
                if catalog_size_col == "None":
                    catalog_size_col = None
                    
            elif catalog_size_type == "Separate value and unit columns":
                size_col_options = ["None"] + catalog_cols
                catalog_size_value_col = st.selectbox(
                    "Size Value Column", 
                    size_col_options,
                    index=size_col_options.index('size_value') if 'size_value' in catalog_cols else 0,
                    key="catalog_size_value"
                )
                catalog_size_unit_col = st.selectbox(
                    "Size Unit Column", 
                    size_col_options,
                    index=size_col_options.index('size_unit') if 'size_unit' in catalog_cols else 0,
                    key="catalog_size_unit"
                )
                if catalog_size_value_col == "None":
                    catalog_size_value_col = None
                if catalog_size_unit_col == "None":
                    catalog_size_unit_col = None
        
        # GTIN Configuration - Only show if GTIN matching is enabled
        catalog_gtin_cols = []
        if enable_gtin_matching:
            st.write("**GTIN/UPC/Barcode Configuration**")
            
            # Smart detection of GTIN columns
            suggested_gtin_cols = smart_detect_gtin_columns(catalog_df)
            if suggested_gtin_cols:
                st.info(f"💡 Detected potential GTIN columns: {', '.join(suggested_gtin_cols)}")
            
            catalog_gtin_cols = st.multiselect(
                "GTIN/UPC/Barcode Columns",
                catalog_cols,
                default=suggested_gtin_cols[:3] if suggested_gtin_cols else [],  # Limit to first 3 suggestions
                key="catalog_gtin",
                help="Select all columns that contain GTIN, UPC, EAN, or barcode data. Multiple columns will be combined for better matching."
            )
            
            if catalog_gtin_cols:
                st.success(f"✅ Selected {len(catalog_gtin_cols)} GTIN column(s)")
            else:
                st.warning("⚠️ No GTIN columns selected - GTIN matching will be disabled for this dataset.")
        
        # Output columns selection
        catalog_output_cols = st.multiselect(
            "Additional Columns to Include in Results",
            catalog_cols,
            default=[],
            key="catalog_output"
        )
        
        # Return same configuration for both "catalog" and "customer" in within-file mode
        return {
            'catalog': {
                'product_cols': catalog_product_cols,
                'product': catalog_product_cols[0] if catalog_product_cols else None, # Primary display column
                'size': catalog_size_col,
                'size_value': catalog_size_value_col,
                'size_unit': catalog_size_unit_col,
                'gtin_cols': catalog_gtin_cols,
                'output_cols': catalog_output_cols
            },
            'customer': {
                'product_cols': catalog_product_cols,
                'product': catalog_product_cols[0] if catalog_product_cols else None, # Primary display column
                'size': catalog_size_col,
                'size_value': catalog_size_value_col,
                'size_unit': catalog_size_unit_col,
                'gtin_cols': catalog_gtin_cols,
                'output_cols': catalog_output_cols
            }
        }
    
    # Original two-file mode
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📦 Catalog File Columns")
        catalog_cols = list(catalog_df.columns)
        
        # Product Name Columns (Multi-select)
        st.write("**Product Name / Description**")
        suggested_product_cols = smart_detect_product_name_columns(catalog_df)
        if suggested_product_cols:
            st.info(f"💡 Detected potential name/description columns: {', '.join(suggested_product_cols)}")
        
        catalog_product_cols = st.multiselect(
            "Product Name / Description Columns (Required)",
            catalog_cols,
            default=suggested_product_cols,
            key="catalog_product_cols",
            help="Select all columns containing product names or descriptions. They will be combined for matching."
        )
        
        if not catalog_product_cols:
            st.warning("⚠️ Please select at least one product name/description column.")
        
        # Size Configuration - Only show if size matching is enabled
        catalog_size_col = None
        catalog_size_value_col = None
        catalog_size_unit_col = None
        
        if include_size_matching:  # Only show size configuration if size matching is enabled
            st.write("**Size Configuration**")
            catalog_size_type = st.radio(
                "How is size data stored?",
                ["No size data", "Combined size column", "Separate value and unit columns"],
                key="catalog_size_type"
            )
            
            if catalog_size_type == "Combined size column":
                size_col_options = ["None"] + catalog_cols
                catalog_size_col = st.selectbox(
                    "Size Column", 
                    size_col_options,
                    index=size_col_options.index('size') if 'size' in catalog_cols else 0,
                    key="catalog_size"
                )
                if catalog_size_col == "None":
                    catalog_size_col = None
                    
            elif catalog_size_type == "Separate value and unit columns":
                size_col_options = ["None"] + catalog_cols
                catalog_size_value_col = st.selectbox(
                    "Size Value Column", 
                    size_col_options,
                    index=size_col_options.index('size_value') if 'size_value' in catalog_cols else 0,
                    key="catalog_size_value"
                )
                catalog_size_unit_col = st.selectbox(
                    "Size Unit Column", 
                    size_col_options,
                    index=size_col_options.index('size_unit') if 'size_unit' in catalog_cols else 0,
                    key="catalog_size_unit"
                )
                if catalog_size_value_col == "None":
                    catalog_size_value_col = None
                if catalog_size_unit_col == "None":
                    catalog_size_unit_col = None
        
        # GTIN Configuration - Only show if GTIN matching is enabled
        catalog_gtin_cols = []
        if enable_gtin_matching:
            st.write("**GTIN/UPC/Barcode Configuration**")
            
            # Smart detection of GTIN columns
            suggested_gtin_cols = smart_detect_gtin_columns(catalog_df)
            if suggested_gtin_cols:
                st.info(f"💡 Detected potential GTIN columns: {', '.join(suggested_gtin_cols)}")
            
            catalog_gtin_cols = st.multiselect(
                "GTIN/UPC/Barcode Columns",
                catalog_cols,
                default=suggested_gtin_cols[:3] if suggested_gtin_cols else [],  # Limit to first 3 suggestions
                key="catalog_gtin",
                help="Select all columns that contain GTIN, UPC, EAN, or barcode data. Multiple columns will be combined for better matching."
            )
            
            if catalog_gtin_cols:
                st.success(f"✅ Selected {len(catalog_gtin_cols)} GTIN column(s)")
            else:
                st.warning("⚠️ No GTIN columns selected - GTIN matching will be disabled for this dataset.")
        
        # Output columns selection
        catalog_output_cols = st.multiselect(
            "Additional Catalog Columns to Include in Results",
            catalog_cols,
            default=[],
            key="catalog_output"
        )
    
    with col2:
        st.subheader("👤 Customer File Columns")
        customer_cols = list(customer_df.columns)
        
        # Product Name Columns (Multi-select)
        st.write("**Product Name / Description**")
        suggested_product_cols = smart_detect_product_name_columns(customer_df)
        if suggested_product_cols:
            st.info(f"💡 Detected potential name/description columns: {', '.join(suggested_product_cols)}")
            
        customer_product_cols = st.multiselect(
            "Product Name / Description Columns (Required)",
            customer_cols,
            default=suggested_product_cols,
            key="customer_product_cols",
            help="Select all columns containing product names or descriptions. They will be combined for matching."
        )
        
        if not customer_product_cols:
            st.warning("⚠️ Please select at least one product name/description column.")
        
        # Size Configuration - Only show if size matching is enabled
        customer_size_col = None
        customer_size_value_col = None
        customer_size_unit_col = None
        
        if include_size_matching:  # Only show size configuration if size matching is enabled
            st.write("**Size Configuration**")
            customer_size_type = st.radio(
                "How is size data stored?",
                ["No size data", "Combined size column", "Separate value and unit columns"],
                key="customer_size_type"
            )
            
            if customer_size_type == "Combined size column":
                customer_size_col_options = ["None"] + customer_cols
                customer_size_col = st.selectbox(
                    "Size Column", 
                    customer_size_col_options,
                    index=customer_size_col_options.index('size') if 'size' in customer_cols else 0,
                    key="customer_size"
                )
                if customer_size_col == "None":
                    customer_size_col = None
                    
            elif customer_size_type == "Separate value and unit columns":
                customer_size_col_options = ["None"] + customer_cols
                customer_size_value_col = st.selectbox(
                    "Size Value Column", 
                    customer_size_col_options,
                    index=customer_size_col_options.index('size_value') if 'size_value' in customer_cols else 0,
                    key="customer_size_value"
                )
                customer_size_unit_col = st.selectbox(
                    "Size Unit Column", 
                    customer_size_col_options,
                    index=customer_size_col_options.index('size_unit') if 'size_unit' in customer_cols else 0,
                    key="customer_size_unit"
                )
                if customer_size_value_col == "None":
                    customer_size_value_col = None
                if customer_size_unit_col == "None":
                    customer_size_unit_col = None
        
        # GTIN Configuration - Only show if GTIN matching is enabled
        customer_gtin_cols = []
        if enable_gtin_matching:
            st.write("**GTIN/UPC/Barcode Configuration**")
            
            # Smart detection of GTIN columns
            suggested_gtin_cols = smart_detect_gtin_columns(customer_df)
            if suggested_gtin_cols:
                st.info(f"💡 Detected potential GTIN columns: {', '.join(suggested_gtin_cols)}")
            
            customer_gtin_cols = st.multiselect(
                "GTIN/UPC/Barcode Columns",
                customer_cols,
                default=suggested_gtin_cols[:3] if suggested_gtin_cols else [],  # Limit to first 3 suggestions
                key="customer_gtin",
                help="Select all columns that contain GTIN, UPC, EAN, or barcode data. Multiple columns will be combined for better matching."
            )
            
            if customer_gtin_cols:
                st.success(f"✅ Selected {len(customer_gtin_cols)} GTIN column(s)")
            else:
                st.warning("⚠️ No GTIN columns selected - GTIN matching will be disabled for this dataset.")
        
        # Output columns selection
        customer_output_cols = st.multiselect(
            "Additional Customer Columns to Include in Results",
            customer_cols,
            default=[],
            key="customer_output"
        )
    
    return {
        'catalog': {
            'product_cols': catalog_product_cols,
            'product': catalog_product_cols[0] if catalog_product_cols else None, # Primary display column
            'size': catalog_size_col,
            'size_value': catalog_size_value_col,
            'size_unit': catalog_size_unit_col,
            'gtin_cols': catalog_gtin_cols,
            'output_cols': catalog_output_cols
        },
        'customer': {
            'product_cols': customer_product_cols,
            'product': customer_product_cols[0] if customer_product_cols else None, # Primary display column
            'size': customer_size_col,
            'size_value': customer_size_value_col,
            'size_unit': customer_size_unit_col,
            'gtin_cols': customer_gtin_cols,
            'output_cols': customer_output_cols
        }
    }
