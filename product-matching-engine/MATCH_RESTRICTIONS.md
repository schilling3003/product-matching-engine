# Match Restrictions Feature

## Overview
The match restrictions feature allows you to limit product matching to only occur between products that share the same values in selected columns. This is particularly useful when you want to:
- Only match products within the same category (e.g., fruits with fruits, vegetables with vegetables)
- Restrict matches to the same brand, department, or commodity type
- Prevent cross-category matches that might be similar but shouldn't be grouped together

## How to Use

1. **Select "Find Similar Within File" mode** - This feature is only available for single-file matching

2. **Enable "Restrict matches to same category"** - Check this box in the sidebar under "Match Restrictions"

3. **Select restriction columns** - Choose one or more columns that products must match on:
   - Common restriction columns include: Category, Commodity, Department, Brand, Type, etc.
   - The system will automatically detect potential restriction columns
   - You can select multiple columns - all must match for products to be considered similar

4. **Run matching as usual** - Products will only match if they meet both:
   - The similarity threshold requirements
   - Have identical values in all selected restriction columns

## Example

With the test data file:
```
Product,Category,Description
Apple,Fresh Fruit,Red delicious apples
Carrot,Vegetable,Orange carrots
Apple,Fresh Fruit,Green apples
```

- Without restrictions: Apple might match Carrot if they have similar text descriptions
- With Category restriction: Apple will only match other products in "Fresh Fruit" category

## Technical Details

- Restriction matching is case-insensitive
- Empty/NaN values are treated as empty strings
- Multiple restriction columns use AND logic (all must match)
- Feature works with both regular and memory-efficient processing
- Restriction status is shown in the sidebar info display

## Tips

- Use columns with limited unique values (2-50 distinct values) for best results
- Avoid using columns with too many unique values (like Product ID)
- Consider using Category, Department, or Type columns for meaningful restrictions
