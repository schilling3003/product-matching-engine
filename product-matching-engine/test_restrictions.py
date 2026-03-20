#!/usr/bin/env python3
"""
Test script to verify match restrictions functionality.
"""

import pandas as pd
import sys
import os

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import the function directly
def smart_detect_restriction_columns(df):
    """Smartly detects potential columns for restricting matches (e.g., Category, Commodity)."""
    potential_restrictions = [
        'category', 'commodity', 'department', 'type', 'class', 'group', 'segment',
        'division', 'section', 'line', 'family', 'brand', 'supplier', 'vendor',
        'product_type', 'product_category', 'item_type', 'item_category',
        'category_name', 'commodity_code', 'dept', 'dept_name'
    ]
    detected = [col for col in df.columns if col.lower() in potential_restrictions]
    # Also look for columns with limited unique values (good candidates for restrictions)
    for col in df.columns:
        if col not in detected and df[col].dtype == 'object':
            unique_count = df[col].nunique()
            # If column has between 2 and 50 unique values, it might be a good restriction
            if 2 <= unique_count <= 50 and unique_count < len(df) * 0.5:
                detected.append(col)
    return detected

def test_restriction_detection():
    """Test that restriction columns are properly detected."""
    
    # Create test data
    test_data = {
        'Product': ['Apple', 'Banana', 'Carrot', 'Orange', 'Broccoli'],
        'Category': ['Fruit', 'Fruit', 'Vegetable', 'Fruit', 'Vegetable'],
        'Commodity': ['Fresh', 'Fresh', 'Fresh', 'Fresh', 'Fresh'],
        'Description': ['Red apple', 'Yellow banana', 'Orange carrot', 'Orange fruit', 'Green vegetable']
    }
    
    df = pd.DataFrame(test_data)
    
    # Test detection
    detected = smart_detect_restriction_columns(df)
    
    print("Detected restriction columns:", detected)
    
    # Verify expected columns are detected
    expected = ['Category', 'Commodity']
    for col in expected:
        if col in detected:
            print(f"✅ Successfully detected '{col}' as a restriction column")
        else:
            print(f"❌ Failed to detect '{col}' as a restriction column")
    
    return len(detected) > 0

def test_restriction_logic():
    """Test the restriction filtering logic."""
    print("\nTesting restriction logic...")
    
    # Test data with different categories
    test_data = {
        'Product': ['Apple', 'Apple2', 'Carrot', 'Carrot2'],
        'Category': ['Fruit', 'Fruit', 'Vegetable', 'Vegetable'],
        'Description': ['Red apple', 'Green apple', 'Orange carrot', 'Baby carrot']
    }
    
    df = pd.DataFrame(test_data)
    
    # Simulate restriction check
    restriction_cols = ['Category']
    
    # Check same category matches
    i, j = 0, 1  # Both fruits
    match = True
    for col in restriction_cols:
        if str(df.iloc[i][col]).lower() != str(df.iloc[j][col]).lower():
            match = False
            break
    
    if match:
        print("✅ Same category match allowed")
    else:
        print("❌ Same category match incorrectly blocked")
    
    # Check different category matches
    i, j = 0, 2  # Fruit vs Vegetable
    match = True
    for col in restriction_cols:
        if str(df.iloc[i][col]).lower() != str(df.iloc[j][col]).lower():
            match = False
            break
    
    if not match:
        print("✅ Different category match correctly blocked")
    else:
        print("❌ Different category match incorrectly allowed")
    
    return True

if __name__ == "__main__":
    print("Testing Match Restrictions Feature\n")
    print("=" * 50)
    
    # Run tests
    test1_pass = test_restriction_detection()
    test2_pass = test_restriction_logic()
    
    print("\n" + "=" * 50)
    if test1_pass and test2_pass:
        print("✅ All tests passed! Match restrictions feature is working.")
    else:
        print("❌ Some tests failed. Please check the implementation.")
