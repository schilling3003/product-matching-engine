#!/usr/bin/env python3
"""
Test script for the new size matching functionality
"""

import sys
import os
sys.path.append(os.path.dirname(__file__))

from app import calculate_size_similarity

def test_size_similarity():
    """Test various size similarity scenarios"""
    
    print("🧪 Testing Size Similarity Function")
    print("=" * 50)
    
    # Test cases: (size1, size2, tolerance, expected_description)
    test_cases = [
        # Exact matches
        ("100.0g", "100.0g", 20, "Exact match should score 100"),
        ("355.0g", "355.0g", 20, "Exact match should score 100"),
        
        # Close matches within tolerance
        ("100.0g", "110.0g", 20, "10% difference, within 20% tolerance"),
        ("100.0g", "90.0g", 20, "10% difference, within 20% tolerance"),
        ("355.0g", "340.0g", 20, "~4% difference, within 20% tolerance"),
        
        # Edge cases at tolerance boundary
        ("100.0g", "120.0g", 20, "20% difference, at tolerance boundary"),
        ("100.0g", "80.0g", 20, "20% difference, at tolerance boundary"),
        
        # Outside tolerance
        ("100.0g", "130.0g", 20, "30% difference, outside 20% tolerance"),
        ("100.0g", "70.0g", 20, "30% difference, outside 20% tolerance"),
        
        # Different tolerance levels
        ("100.0g", "110.0g", 5, "10% difference, outside 5% tolerance"),
        ("100.0g", "105.0g", 5, "5% difference, at 5% tolerance boundary"),
        ("100.0g", "103.0g", 5, "3% difference, within 5% tolerance"),
        
        # Empty/invalid cases
        ("", "100.0g", 20, "Empty size should score 0"),
        ("100.0g", "", 20, "Empty size should score 0"),
        ("", "", 20, "Both empty should score 0"),
        ("invalid", "100.0g", 20, "Invalid format should score 0"),
        
        # Real-world examples
        ("355.0g", "330.0g", 20, "12 fl oz vs ~11.6 fl oz cans"),
        ("473.2g", "500.0g", 20, "16 fl oz vs ~17.6 fl oz bottles"),
        ("28.4g", "30.0g", 20, "1 oz vs ~1.06 oz packages"),
    ]
    
    print(f"{'Size 1':<10} {'Size 2':<10} {'Tolerance':<9} {'Score':<6} {'Description'}")
    print("-" * 80)
    
    for size1, size2, tolerance, description in test_cases:
        score = calculate_size_similarity(size1, size2, tolerance)
        print(f"{size1:<10} {size2:<10} {tolerance:<9}% {score:<6.1f} {description}")
    
    print("\n" + "=" * 50)
    print("✅ Size similarity testing complete!")
    
    # Test some specific scenarios
    print("\n🔍 Detailed Analysis Examples:")
    
    # Example 1: Cola cans
    score1 = calculate_size_similarity("355.0g", "330.0g", 20)  # 12 fl oz vs 11.6 fl oz
    print(f"Cola cans (355ml vs 330ml): {score1:.1f}% similarity")
    
    # Example 2: Different bottle sizes
    score2 = calculate_size_similarity("500.0g", "473.2g", 20)  # 500ml vs 16 fl oz
    print(f"Bottles (500ml vs 16 fl oz): {score2:.1f}% similarity")
    
    # Example 3: Small packages
    score3 = calculate_size_similarity("28.4g", "25.0g", 20)  # 1 oz vs ~0.88 oz
    print(f"Small packages (1 oz vs 25g): {score3:.1f}% similarity")

if __name__ == "__main__":
    test_size_similarity()
