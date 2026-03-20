"""
GTIN Processing Module

This module handles the normalization and variant generation pipeline for GTIN matching.
It can handle inconsistent and messy GTIN data, including missing check digits,
missing leading zeros, and mixed formats.
"""

import re
import pandas as pd
import numpy as np
from typing import List, Set, Tuple, Optional, Dict


def calculate_gtin_check_digit(gtin_base: str) -> str:
    """
    Calculate the check digit for a GTIN.
    
    Args:
        gtin_base (str): GTIN without check digit (7, 11, or 13 digits)
    
    Returns:
        str: Complete GTIN with check digit
    """
    if not gtin_base or not gtin_base.isdigit():
        return ""
    
    # Pad to appropriate length for check digit calculation
    if len(gtin_base) == 7:
        # GTIN-8: pad to 7 digits for calculation
        padded = gtin_base.zfill(7)
    elif len(gtin_base) == 11:
        # GTIN-12: pad to 11 digits for calculation  
        padded = gtin_base.zfill(11)
    elif len(gtin_base) == 13:
        # GTIN-14: pad to 13 digits for calculation
        padded = gtin_base.zfill(13)
    else:
        return ""
    
    # Calculate check digit using GTIN algorithm
    total = 0
    for i, digit in enumerate(reversed(padded)):
        weight = 3 if i % 2 == 0 else 1
        total += int(digit) * weight
    
    check_digit = (10 - (total % 10)) % 10
    return padded + str(check_digit)


def validate_gtin_check_digit(gtin: str) -> bool:
    """
    Validate if a GTIN has the correct check digit.
    
    Args:
        gtin (str): Complete GTIN to validate
    
    Returns:
        bool: True if check digit is valid
    """
    if not gtin or not gtin.isdigit() or len(gtin) not in [8, 12, 13, 14]:
        return False
    
    # Extract base and check digit
    base = gtin[:-1]
    expected_check = gtin[-1]
    
    # Calculate what the check digit should be
    calculated_gtin = calculate_gtin_check_digit(base)
    
    return calculated_gtin == gtin


def correct_gtin_check_digit(gtin: str) -> str:
    """
    Correct the check digit of a GTIN if it's invalid.
    
    Args:
        gtin (str): GTIN that may have incorrect check digit
    
    Returns:
        str: GTIN with corrected check digit, or empty string if invalid
    """
    if not gtin or not gtin.isdigit() or len(gtin) not in [8, 12, 13, 14]:
        return ""
    
    base = gtin[:-1]
    return calculate_gtin_check_digit(base)


def extract_unit_gtin_from_case(case_gtin: str) -> str:
    """
    Extract the unit GTIN from a GTIN-14 case code.
    GTIN-14 case codes typically have the format: [indicator][unit_gtin][check]
    
    Args:
        case_gtin (str): 14-digit case GTIN
    
    Returns:
        str: 13-digit unit GTIN, or empty string if not applicable
    """
    if not case_gtin or len(case_gtin) != 14 or not case_gtin.isdigit():
        return ""
    
    # Extract the unit portion (remove first digit indicator)
    unit_base = case_gtin[1:13]  # Remove indicator digit and check digit
    
    # Calculate proper check digit for the unit GTIN
    unit_gtin = calculate_gtin_check_digit(unit_base)
    
    return unit_gtin if len(unit_gtin) == 13 else ""


def normalize_and_generate_variants(gtin_value: str) -> Dict[str, str]:
    """
    Core normalization and variant generation pipeline.
    Processes a raw GTIN value and generates all possible valid interpretations.
    
    Args:
        gtin_value (str): Raw GTIN value from data
    
    Returns:
        Dict[str, str]: Dictionary mapping GTIN-14 variants to their source type
                       ('original', 'corrected', 'case_to_unit', 'missing_check')
    """
    if not gtin_value:
        return {}
    
    # Clean input - remove all non-numeric characters
    cleaned = re.sub(r'[^\d]', '', str(gtin_value))
    
    if not cleaned:
        return {}
    
    variants = {}  # Dict mapping GTIN to source type
    
    # Generate variants based on length - be more aggressive about variant generation
    length = len(cleaned)
    
    # Always add the original (padded to 14 digits)
    if length <= 14:
        padded = cleaned.zfill(14)
        variants[padded] = 'original'
    
    # Try to generate corrected version
    if length in [8, 12, 13, 14]:
        corrected = correct_gtin_check_digit(cleaned)
        if corrected and corrected != cleaned:
            variants[corrected.zfill(14)] = 'corrected'
    
    # Try to complete incomplete GTINs
    if length in [7, 11]:
        complete_gtin = calculate_gtin_check_digit(cleaned)
        if complete_gtin:
            variants[complete_gtin.zfill(14)] = 'missing_check'
    
    # For 12-digit, try both UPC-A and EAN-13 interpretations
    if length == 12:
        # As EAN-13 base (add check digit)
        ean13_complete = calculate_gtin_check_digit(cleaned)
        if ean13_complete and ean13_complete != cleaned:
            variants[ean13_complete.zfill(14)] = 'missing_check'
    
    # For 14-digit, try case-to-unit extraction
    if length == 14:
        unit_gtin = extract_unit_gtin_from_case(cleaned)
        if unit_gtin:
            variants[unit_gtin.zfill(14)] = 'case_to_unit'
            # Also try corrected unit GTIN
            unit_corrected = correct_gtin_check_digit(unit_gtin)
            if unit_corrected and unit_corrected != unit_gtin:
                variants[unit_corrected.zfill(14)] = 'case_to_unit'
    
    # Generate additional variants by trying different interpretations
    # This is more aggressive to match original behavior
    for i in range(max(0, 14 - length)):
        alt_padded = ('0' * i) + cleaned + ('0' * (14 - length - i))
        if len(alt_padded) == 14 and alt_padded not in variants:
            variants[alt_padded] = 'original'
    
    # Remove any empty or invalid strings
    variants = {k: v for k, v in variants.items() if k and len(k) == 14 and k.isdigit()}
    
    return variants


def consolidate_gtin_columns(df: pd.DataFrame, gtin_columns: List[str]) -> pd.Series:
    """
    Consolidate multiple GTIN columns into a single series of GTIN pools.
    
    Args:
        df (pd.DataFrame): DataFrame containing GTIN columns
        gtin_columns (List[str]): List of column names containing GTINs
    
    Returns:
        pd.Series: Series where each element is a dict of normalized GTIN-14 variants to source types
    """
    if not gtin_columns or df.empty:
        return pd.Series([{}] * len(df))
    
    def consolidate_row_gtins(row):
        all_variants = {}
        for col in gtin_columns:
            if col in df.columns and pd.notna(row[col]):
                variants = normalize_and_generate_variants(str(row[col]))
                all_variants.update(variants)
        return all_variants
    
    return df.apply(consolidate_row_gtins, axis=1)


def calculate_gtin_match_confidence(customer_gtins: Dict[str, str], catalog_gtins: Dict[str, str]) -> Tuple[float, str, List[str]]:
    """
    Calculate the confidence level and details of a GTIN match.
    
    Args:
        customer_gtins (Dict[str, str]): Dict of customer GTIN variants to source types
        catalog_gtins (Dict[str, str]): Dict of catalog GTIN variants to source types
    
    Returns:
        Tuple[float, str, List[str]]: (confidence_score, match_type, matching_gtins)
    """
    if not customer_gtins or not catalog_gtins:
        return 0.0, "No Match", []
    
    # Find intersection of GTINs
    customer_gtin_set = set(customer_gtins.keys())
    catalog_gtin_set = set(catalog_gtins.keys())
    matching_gtins = list(customer_gtin_set.intersection(catalog_gtin_set))
    
    if not matching_gtins:
        return 0.0, "No Match", []
    
    # Determine the highest confidence match type
    best_confidence = 0.0
    best_match_type = "No Match"
    
    for gtin in matching_gtins:
        customer_source = customer_gtins[gtin]
        catalog_source = catalog_gtins[gtin]
        
        # Determine confidence based on source types
        if customer_source == 'original' and catalog_source == 'original':
            confidence = 120.0  # Exact Match
            match_type = "Exact Match"
        elif 'corrected' in [customer_source, catalog_source]:
            confidence = 92.0   # Corrected Match
            match_type = "Corrected Match"
        elif 'case_to_unit' in [customer_source, catalog_source]:
            confidence = 90.0   # Case/Unit Match
            match_type = "Case/Unit Match"
        elif 'missing_check' in [customer_source, catalog_source]:
            confidence = 92.0   # Missing check digit (treat as corrected)
            match_type = "Corrected Match"
        else:
            confidence = 120.0  # Default to exact match
            match_type = "Exact Match"
        
        # Keep the highest confidence match
        if confidence > best_confidence:
            best_confidence = confidence
            best_match_type = match_type
    
    return best_confidence, best_match_type, matching_gtins


def smart_detect_gtin_columns(df: pd.DataFrame) -> List[str]:
    """
    Automatically detect columns that likely contain GTIN/UPC/Barcode data.
    
    Args:
        df (pd.DataFrame): DataFrame to analyze
    
    Returns:
        List[str]: List of column names that likely contain GTINs
    """
    if df.empty:
        return []
    
    gtin_patterns = [
        r'gtin', r'upc', r'ean', r'barcode', r'bar_code', r'product_code',
        r'item_code', r'sku', r'code', r'id'
    ]
    
    # Exclude patterns that are likely not GTINs
    exclude_patterns = [r'type', r'description', r'name', r'category']
    
    potential_columns = []
    
    for col in df.columns:
        col_lower = col.lower()
        
        # Check if column name matches GTIN patterns
        matches_gtin = any(re.search(pattern, col_lower) for pattern in gtin_patterns)
        matches_exclude = any(re.search(pattern, col_lower) for pattern in exclude_patterns)
        
        if matches_gtin and not matches_exclude:
            # Additional validation: check if column contains numeric-like data
            sample_values = df[col].dropna().head(10)
            if not sample_values.empty:
                # Check if most values are numeric or mostly numeric
                numeric_count = 0
                for val in sample_values:
                    cleaned_val = re.sub(r'[^\d]', '', str(val))
                    if cleaned_val and len(cleaned_val) >= 7:  # Minimum reasonable GTIN length
                        numeric_count += 1
                
                if numeric_count >= len(sample_values) * 0.5:  # At least 50% look like GTINs
                    potential_columns.append(col)
    
    return potential_columns


def generate_gtin_quality_report(df: pd.DataFrame, gtin_columns: List[str]) -> Dict:
    """
    Generate a quality report for GTIN data in the dataset.
    
    Args:
        df (pd.DataFrame): DataFrame containing GTIN data
        gtin_columns (List[str]): List of GTIN column names
    
    Returns:
        Dict: Quality report with statistics
    """
    if not gtin_columns or df.empty:
        return {
            'total_products': 0,
            'products_with_gtins': 0,
            'valid_gtins': 0,
            'correctable_gtins': 0,
            'invalid_gtins': 0,
            'coverage_percentage': 0.0
        }
    
    total_products = len(df)
    products_with_gtins = 0
    valid_gtins = 0
    correctable_gtins = 0
    invalid_gtins = 0
    
    for _, row in df.iterrows():
        has_gtin = False
        for col in gtin_columns:
            if col in df.columns and pd.notna(row[col]):
                gtin_value = str(row[col])
                cleaned = re.sub(r'[^\d]', '', gtin_value)
                
                if cleaned:
                    has_gtin = True
                    
                    # Check if it's a valid length
                    if len(cleaned) in [7, 8, 11, 12, 13, 14]:
                        if len(cleaned) in [8, 12, 13, 14]:
                            # Check if check digit is valid
                            if validate_gtin_check_digit(cleaned):
                                valid_gtins += 1
                            else:
                                # Check if it's correctable
                                corrected = correct_gtin_check_digit(cleaned)
                                if corrected:
                                    correctable_gtins += 1
                                else:
                                    invalid_gtins += 1
                        else:
                            # Missing check digit - correctable
                            correctable_gtins += 1
                    else:
                        invalid_gtins += 1
        
        if has_gtin:
            products_with_gtins += 1
    
    coverage_percentage = (products_with_gtins / total_products * 100) if total_products > 0 else 0
    
    return {
        'total_products': total_products,
        'products_with_gtins': products_with_gtins,
        'valid_gtins': valid_gtins,
        'correctable_gtins': correctable_gtins,
        'invalid_gtins': invalid_gtins,
        'coverage_percentage': coverage_percentage
    }
